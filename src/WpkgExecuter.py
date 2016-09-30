# -*- coding: utf-8 -*-
"""WPKGExecuter.py
Class for executing WPKG
"""
import WpkgConfig
import WpkgWriter
import WpkgNetworkHandler
import WpkgOutputParser
import WpkgRebootHandler
import logging
import sys, os, re, subprocess, time

try:
    from Queue import Queue, Empty
except ImportError:
    from queue import Queue, Empty  # python 3.x
from threading import Thread


def enqueue_output(out, queue):
    for line in iter(out.readline, ''):
        queue.put(line)
    out.close()


class NullHandler(logging.Handler):
    def emit(self, record):
        pass


class WpkgExecuter():
    
    is_running = False
    
    def __init__(self, handle=None):
        self.config = WpkgConfig.WpkgConfig()
        self.wpkg_command = self.config.get("WpkgCommand")
        self.writer = WpkgWriter.WpkgWriter(handle)
        self.network_handler = WpkgNetworkHandler.WpkgNetworkHandler()
        self.parser = WpkgOutputParser.WpkgOutputParser()
        self.reboot_handler = WpkgRebootHandler.WpkgRebootHandler()
        self.parse_wpkg_command()
        self.activityvalue = 0
    
    def parse_wpkg_command(self):
        commandstring = self.wpkg_command
        commandstring = os.path.expandvars(commandstring) #Expanding variables

        # check if starts with cscript and contains /noreboot /synchronize
        # and/or /sendStatus is in command. If not, add it.
        
        # split command line by space except when quoted - preserve quotes
        commandlist = re.findall(r'(?:[^\s"]|"(?:\\.|[^"])*")+', commandstring)
        is_js_script = False
        
        # remove possible quotes before checking whether it is a js script
        jscommand = re.sub(r'^"|"$', '', commandlist[0]).lower()

        if jscommand == "cscript" or jscommand[-3:]==".js":
           is_js_script = True
        if is_js_script == True:
            if commandlist[0].lower() != "cscript":
                logger.debug("WpkgCommand is a js file but is missing 'cscript', adding")
                commandlist.insert(0, "cscript")
            if not "/noreboot" in commandlist:
                logger.debug("WpkgCommand is a js but is missing /noreboot, adding")
                commandlist.append("/noreboot")
            if not "/synchronize" in commandlist:
                logger.debug("WpkgCommand is a js but is missing /synchronize, adding")
                commandlist.append("/synchronize")
            if not "/sendStatus" in commandlist:
                logger.debug("WpkgCommand is a js but is missing /sendStatus, adding")
                commandlist.append("/sendStatus")
            if not "/nonotify" in commandlist:
                logger.debug("WpkgCommand is a js but is missing /nonotify, adding")
                commandlist.append("/nonotify")
            if not "/quiet" in commandlist:
                logger.debug("WpkgCommand is a js but is missing /quiet, adding")
                commandlist.append("/quiet")
        self.execute_command = " ".join(commandlist)

    def Query(self, handle=None):
        self.writer = WpkgWriter.WpkgWriter(handle)

        # Adding query and dryrun parameter to execute command
        # /dryrun is used to the file date of wpkg.xml is untouched
        self.query_command = self.execute_command + ' /query:Iudr /dryrun'
        if self.is_running:
            logger.info(R"Client requested WPKG to execute query, but WPKG is already running.")
            msg = "201 " + _("Info: WPKG is already running a task.")
            self.writer.Write(msg)
            return

        parsedline = _("Initializing Wpkg-GP software query")
        self.writer.Write("100 " + parsedline)
        logger.info(R"Executing WPKG with the command %s" % self.execute_command)

        # Open the network share as another user, if necessary
        if not self.network_handler.connect_to_network_share():
            net_msg = _("Error: Connecting to network share failed.")
            self.writer.Write("204 " + net_msg)
            logger.error("Connecting to network share failed. Exiting.")
            return

        # Check if System is on Blacklist
        if not self.allowed_to_execute():
            net_msg = _("Info: Client was blocked from server to execute wpkg.")
            self.writer.Write("205 " + net_msg)
            logger.info("Client was blocked from server to execute wpkg.")
            return

        # Add environment parameters
        env = os.environ.copy()
        config_env = self.config.EnvironmentVariables.get()
        if config_env != None:
            env.update(config_env)
        # logger.debug(R"Environment variables are: %s" % env)

        # Set wpkg runningstate true
        self.config.set_wpkg_runningstate('true')

        # Run WPKG Query
        self.proc = subprocess.Popen(self.query_command, stdout=subprocess.PIPE, bufsize=1, universal_newlines=True,
                                     env=env)
        self.isrunning = True

        output = self.proc.communicate()
        lines = output[0].split('\n')
        exitcode = self.proc.poll()
        self.is_running = False

        logger.info(R"Finished executing Wpkg.js Query")

        if exitcode == 1:  # Cscript returned an error
            logger.error(R"WPKG command returned an error: %s" % lines[-1:])
            self.writer.Write("200 " + _("Wpkg returned an error: %s") % lines[-1][0:-1])
            return

        value_dict = {'Installation pending':'install',
                      'Upgrade pending':'update',
                      'Downgrade pending': 'downgrade',
                      'Remove pending': 'remove'}
        excludes = ('ID:', 'Reboot:', 'Execute:', 'Priority:', 'Status:', 'Revision (old):')

        # Remove the 4 leading lines
        [lines.pop(0) for n in [0, 1, 2, 3]]
        cleaned_lines = []

        # filter output more for relevant information
        for line in lines:
            # Remove Leading Spaces
            line = line.lstrip()
            # Remove double spaces
            line = re.sub('\s{2,}', '', line)
            if line != '':
                # If line not empty add data to the cleaned_lines list
                if line.startswith(excludes):
                    # exclude data we don't want
                    continue
                elif line.startswith('Revision:'):
                    value = line.replace('Revision:', '')
                elif line.startswith('Revision (new):'):
                    value = line.replace('Revision (new):', '')
                elif line.startswith('Action:'):
                    value = line.replace('Action:', '')
                    value = value_dict[value]
                else:
                    # Package Name
                    value = line
                cleaned_lines.append(value)

        # Closing handle to share
        self.network_handler.disconnect_from_network_share()

        # convert list into list with sublists (every 3 entries become a sublist)
        composite_list = [cleaned_lines[x:x + 3] for x in range(0, len(cleaned_lines), 3)]
        # TODO: Add Sorting to composite_list?

        if composite_list:
            for name, version, task in composite_list:
                # Write Info to pipe
                query_msg = "103 TASK: %s\tNAME: %s\tREVISION: %s" % (task, name, version)
                self.writer.Write(query_msg)
        else:
            query_msg = "104 " + _("No pending wpkg tasks")
            self.writer.Write(query_msg)

    def Execute(self, handle=None, rebootcancel=False):
        self.writer = WpkgWriter.WpkgWriter(handle)
        lines = []
        if self.is_running:
            logger.info(R"Client requested WPKG to execute, but WPKG is already running.")
            msg = "201 " + _("Info: WPKG is already running a task.")
            self.writer.Write(msg)
            return

        parsedline = _("Initializing Wpkg-GP software installation")
        self.writer.Write("100 " + parsedline)
        logger.info(R"Executing WPKG with the command %s" % self.execute_command)

        #Open the network share as another user, if necessary
        if not self.network_handler.connect_to_network_share():
            net_msg = _("Error: Connecting to network share failed.")
            self.writer.Write("204 " + net_msg)
            logger.error("Connecting to network share failed. Exiting.")
            if not rebootcancel:
                time.sleep(2)
            return

        # Check if System is on Blacklist
        if not self.allowed_to_execute():
            net_msg = _("Info: Client was blocked from server to execute wpkg.")
            self.writer.Write("205 " + net_msg)
            logger.info("Client was blocked from server to execute wpkg.")
            if not rebootcancel:
                # Enough time to see the message during bootup
                time.sleep(4)
            return

        # Add environment parameters
        env = os.environ.copy()
        config_env = self.config.EnvironmentVariables.get()
        if config_env != None:
            env.update(config_env)
        #logger.debug(R"Environment variables are: %s" % env)

        # Set wpkg runningstate true
        self.config.set_wpkg_runningstate('true')

        # Run WPKG
        self.proc = subprocess.Popen(self.execute_command, stdout=subprocess.PIPE, bufsize=1, universal_newlines=True, env=env)
        self.isrunning = True

        q = Queue()
        t = Thread(target=enqueue_output, args=(self.proc.stdout, q))
        t.daemon = True
        t.start()

        if self.config.get("WpkgActivityIndicator") == 1:
            show_activity = True

        #Reading lines
        quit = False
        lastsec = None
        while 1:
            try:
                line = q.get(timeout=0.05)
            except Empty:
                if quit:
                    break # Now we have appended the last line
                if show_activity:
                    currsec = time.time()
                    if(lastsec != None and currsec - lastsec >= 1): #Show every 1 sec
                        self.writer.Write("101 %s%s" % (parsedline, self.GetActivityIndicator()))
                        lastsec = currsec
            else:
                lines.append(line)
                if quit:
                    break # Now we have appended the last line
                self.parser.parse_line(line)
                if self.parser.updated:
                    parsedline = self.parser.get_formatted_line()
                    self.writer.Write("100 %s      " % parsedline)
                    lastsec = time.time() # Reset timer
            if self.proc.poll() != None: #Wpkg is finished
                self.is_running = False
                quit = True # Run a last loop to fetch the last line
            
        self.parser.reset()
        
        exitcode = self.proc.poll()
        #Closing handle to share
        self.network_handler.disconnect_from_network_share()
        logger.info(R"Finished executing Wpkg.js")
            
        if exitcode == 1: #Cscript returned an error
            logger.error(R"WPKG command returned an error: %s" % lines[-1:])
            self.writer.Write("200 " + _("Wpkg returned an error: %s") % lines[-1][0:-1])
            return
        
        if exitcode == 770560: #WPKG returns this when it requests a reboot
            logger.info(R"WPKG requested a reboot")
            status = self.reboot_handler.reboot(rebootcancel)
            self.writer.Write(status)
        else:
            self.reboot_handler.reset_reboot_number()

    def Cancel(self, handle=sys.stdout):
        if self.isrunning:
            self.proc.kill()
            logger.info("Cancel called, WPKG process was killed.")
            msg = "105 " + _("Cancel called, WPKG process was killed")
        else:
            logger.info("Cancel called, but WPKG process was not running")
            msg = "202 " + _("Cancel called, WPKG process was not running")
        try:
            self.writer.Write(handle, msg)
        except TypeError: #Maybe pipe is closed now
            pass
        self.config.set_wpkg_runningstate('false')

    def GetActivityIndicator(self):
        # Show for every 10 iteration
        mod = self.activityvalue % 5
        self.activityvalue = self.activityvalue + 1
        if mod == 0:
            return "...    "
        if mod == 1:
            return " ...   "
        if mod == 2:
            return "  ...  "
        if mod == 3:
            return "   ... "
        if mod == 4:
            return "    ..."

    def allowed_to_execute(self):
        allowed = True
        block_all = False
        commandstring = os.path.expandvars(self.wpkg_command)
        wpkg_path = commandstring.split("wpkg.js", 1)[0]
        blacklist_path = wpkg_path + 'blacklist.txt'
        try:
            with open(blacklist_path, "r") as blacklist_file:
                data = blacklist_file.readlines()
            blacklist = []
            for entry in data:
                entry = entry.replace('\n', '')
                if not entry.startswith('#') and entry != '':
                    blacklist.append(entry.lower().strip())
                if entry.strip().lower() == '!all!':
                    # Block all systems from executing
                    block_all = True
                    break
            hostname = os.getenv('computername').lower()
            # If Hostname in Blacklist don't allow execution
            if hostname in blacklist:
                allowed = False
        except IOError:
            return allowed
        if block_all:
            return False
        else:
            return allowed

if __name__=='__main__':
    import sys, gettext
    gettext.install('wpkg-gp')
    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")                        
    h = logging.StreamHandler(sys.stdout)
    h.setFormatter(formatter)
    logger = logging.getLogger("WpkgExecuter")
    logger.addHandler(h)
    logger.setLevel(logging.DEBUG)
    WPKG = WpkgExecuter()
    WPKG.Execute()
else:
    h = NullHandler()
    logger = logging.getLogger("WpkgService")
    logger.addHandler(h)
