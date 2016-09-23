# Service controlling Wpkg-GP

# A multi-threaded service that executes WPKG.js and prompts
# returns it's output via a named pipe

import win32serviceutil, win32service
import pywintypes, win32con, winerror
from win32event import *
from win32file import *
from win32pipe import *
from win32api import *
from win32security import *
from ntsecuritycon import *
import traceback
import thread
import servicemanager
import WpkgExecuter
import WpkgLGPUpdater
import WpkgTranslator
import WpkgConfig
import _winreg, logging, logging.handlers
import os.path, sys
import gettext

MY_PIPE_NAME = r"\\.\pipe\WPKG"
# From http://msdn.microsoft.com/en-us/library/aa379649%28VS.85%29.aspx
SID_LOCAL = "S-1-2-0"
SID_ADMINISTRATORS = "S-1-5-32-544"


def ApplyIgnoreError(fn, args):
    try:
        return fn(*args)
    except error: # Ignore win32api errors.
    #except pywintypes.api_error:
        return None
        

class WPKGControlService(win32serviceutil.ServiceFramework):
    _svc_name_ = "WpkgServer"
    _svc_display_name_ = "WPKG Control Service"
    _svc_description_ = "Controller service for userspace WPKG management applications. (http://wpkg-gp.googlecode.com/)"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = CreateEvent(None, 0, 0, None)
        self.overlapped = pywintypes.OVERLAPPED()
        self.overlapped.hEvent = CreateEvent(None,0,0,None)
        self.thread_handles = []

        self.config = WpkgConfig.WpkgConfig()

        #reset wpkg runningstate
        self.config.set_wpkg_runningstate('false')

        verbosity = self.config.get("WpkgVerbosity")
        install_path = self.config.install_path
        
        self.logger = logging.getLogger("WpkgService")
        logdir = os.path.join(install_path, "logs")

        try:
            os.makedirs(logdir)
        except WindowsError:
            pass
        logfile = os.path.join(logdir, "WpkgService.log")
        handler = logging.handlers.RotatingFileHandler(logfile, maxBytes=200000, backupCount=2)
        
        if verbosity == 3:
            log_level = logging.DEBUG
        elif verbosity == 2:
            log_level = logging.INFO
        elif verbosity == 1:
            log_level = logging.ERROR
        else:
            log_level = logging.CRITICAL
        
        self.logger.setLevel(log_level)
        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")        
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.logger.info("Logging started with verbosity: %i" % verbosity)

        self.translator = WpkgTranslator.WpkgTranslator()
        self.translator.install()

        # Enable/Disable LGP
        LGP_handler = WpkgLGPUpdater.WpkgLocalGPConfigurator()
        LGP_handler.update()
        
        self.WpkgExecuter = WpkgExecuter.WpkgExecuter()
    
    def CreatePipeSecurityObject(self):
        # Create a security object giving World read/write access,
        # but only "Owner" modify access.
        sa = pywintypes.SECURITY_ATTRIBUTES()
        sidEveryone = pywintypes.SID()
        sidEveryone.Initialize(SECURITY_WORLD_SID_AUTHORITY,1)
        sidEveryone.SetSubAuthority(0, SECURITY_WORLD_RID)

        #sidLocalAdministrator = pywintypes.SID()
        #sidLocalAdministrator.Initialize(SECURITY_NT_AUTHORITY,2)
        #sidLocalAdministrator.SetSubAuthority(0, SECURITY_BUILTIN_DOMAIN_RID)
        #sidLocalAdministrator.SetSubAuthority(1, DOMAIN_ALIAS_RID_ADMINS)
            
        sidCreator = pywintypes.SID()
        sidCreator.Initialize(SECURITY_CREATOR_SID_AUTHORITY,1)
        sidCreator.SetSubAuthority(0, SECURITY_CREATOR_OWNER_RID)

        acl = pywintypes.ACL()
        acl.AddAccessAllowedAce(FILE_GENERIC_READ|FILE_GENERIC_WRITE, sidEveryone)
        #acl.AddAccessAllowedAce(FILE_GENERIC_READ|FILE_GENERIC_WRITE, sidLocalAdministrator)
        acl.AddAccessAllowedAce(FILE_ALL_ACCESS, sidCreator)

        sa.SetSecurityDescriptorDacl(1, acl, 0)
        return sa

    def CheckIfClientIsAllowedToExecute(self, handle):
        # Check security
        self.logger.debug("Checking client acccess")
        ImpersonateNamedPipeClient(handle)
        token = OpenThreadToken(GetCurrentThread(), TOKEN_ALL_ACCESS, 1)
        groups = GetTokenInformation(token, TokenGroups)
        is_local_admin = False
        is_local_user = False
        for group in groups:
            group_sid = group[0]
            try:
                user, domain, type = LookupAccountSid (None, group_sid)
                #self.logger.debug(user)
                administrators_sid = GetBinarySid(SID_ADMINISTRATORS)
                local_sid = GetBinarySid(SID_LOCAL)
                
                if ConvertSidToStringSid(group_sid) == ConvertSidToStringSid(administrators_sid):
                    self.logger.debug("Client is a member of Administrators group")
                    is_local_admin = True
                if ConvertSidToStringSid(group_sid) == ConvertSidToStringSid(local_sid):
                    self.logger.debug("Client is a local user")
                    is_local_user = True
            except error as (n, f, d):
                if n == 1332: # No mapping between account names and security ID
                    pass
                else:
                    RevertToSelf()
                    raise

        RevertToSelf()
        
        execute_by_nonadmins = self.config.get("WpkgExecuteByNonAdmins")
        self.logger.debug("WpkgExecuteByNonAdmins is %i" % execute_by_nonadmins)

        execute_by_local_users = self.config.get("WpkgExecuteByLocalUsers")
        self.logger.debug("WpkgExecuteByLocalUsers is %i" % execute_by_local_users)

        allow_execution = False
        
        if is_local_admin:
            self.logger.debug("Client user is a member of Administrators group, permission is granted")
            allow_execution = True
        elif execute_by_nonadmins == 1:
            self.logger.debug("All users may access the service, persmission is granted")
            allow_execution = True
        elif execute_by_local_users == 1 and is_local_user:
            self.logger.debug("Client user is local user, permission is granted")
            allow_execution = True
        else:
            self.logger.debug("Permission to execute is not given.")
            allow_execution = False
        return allow_execution
    
    def DoProcessClient(self, pipeHandle, tid):
        self.logger.debug("DoProcessClient() start")
        rebootcancel = False
        try:
            try:
                # Create a loop, reading large data.  If we knew the data stream was
                # was small, a simple ReadFile would do.
                d = ''.encode('ascii') # ensure bytes on py2k and py3k...
                hr = winerror.ERROR_MORE_DATA
                while hr==winerror.ERROR_MORE_DATA:
                    hr, thisd = ReadFile(pipeHandle, 256)
                    d = d + thisd
                    d = d.rstrip("\0") #remove trailing nulls
                ok = 1
            except error:
                self.logger.info("Client disconnected")
                # Client disconnection - do nothing
                ok = 0

            # A secure service would handle (and ignore!) errors writing to the
            # pipe
            if ok:
                if self.WpkgExecuter.is_running:
                    msg = "200 " + self.WpkgExecuter.getStatus()
                    self.logger.info("Wpkg Executer is not ready. Returning '%s' to client." % msg)
                    WriteFile(pipeHandle, msg.encode('ascii'))
                else:
                    if d == b"Execute" or d == b"ExecuteFromGPE" or d == b"ExecuteNoReboot":
                        if d == b"ExecuteNoReboot":
                            rebootcancel = True
                        if d == b"ExecuteFromGPE" and self.config.get("DisableAtBootUp") == 1:
                            self.logger.info("Excution at startup is disabled, will not run".encode('ascii'))
                            WriteFile(pipeHandle, "200 Excution at startup is disabled, will not run")
                        else:
                            self.logger.info("Received 'Execute', executing WPKG")
                            if self.CheckIfClientIsAllowedToExecute(pipeHandle):
                                self.WpkgExecuter.Execute(handle=pipeHandle, rebootcancel=rebootcancel)
                            else:
                                self.logger.info("The user trying to execute Wpkg-GP is not authorized to do so")
                                WriteFile(pipeHandle, "200 Info: You are not authorized to execute Wpkg-GP".encode('ascii'))
                    elif d == b"Query":
                        self.logger.info("Received 'Query', querying WPKG for updates")
                        if self.CheckIfClientIsAllowedToExecute(pipeHandle):
                            self.WpkgExecuter.Query(handle=pipeHandle)
                        else:
                            self.logger.info("The user trying to execute Wpkg-GP is not authorized to do so")
                            WriteFile(pipeHandle, "200 Info: You are not authorized to execute Wpkg-GP".encode('ascii'))
                    elif d == b"Cancel":
                        self.logger.info("Received 'Cancel', cancelling WPKG")
                        if self.CheckIfClientIsAllowedToExecute(pipeHandle):
                            self.WpkgExecuter.Cancel(pipeHandle)
                        else:
                            self.logger.info("The user trying to execute Wpkg-GP is not authorized to do so")
                            WriteFile(pipeHandle, "200 Info: You are not authorized to execute Wpkg-GP".encode('ascii'))
                    else:
                        msg = "203 Unknown command: %s" % d
                        self.logger.info("Sending '%s' to client" % msg)
                        WriteFile(pipeHandle, msg.encode('ascii'))

                #msg = ("%s (on thread %d) sent me %s" % (GetNamedPipeHandleState(pipeHandle)[4],tid, d)).encode('ascii')
                #WriteFile(pipeHandle, msg)
        except Exception, e:
            self.logger.exception("Error when processing Named Pipe Client:")
            raise
        finally:
            ApplyIgnoreError( DisconnectNamedPipe, (pipeHandle,) )
            ApplyIgnoreError( CloseHandle, (pipeHandle,) )

    def ProcessClient(self, pipeHandle):
        try:
            procHandle = GetCurrentProcess()
            th = DuplicateHandle(procHandle, GetCurrentThread(), procHandle, 0, 0, win32con.DUPLICATE_SAME_ACCESS)
            try:
                self.thread_handles.append(th)
                try:
                    return self.DoProcessClient(pipeHandle, th)
                except:
                    traceback.print_exc()
            finally:
                self.thread_handles.remove(th)
        except:
            traceback.print_exc()

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        # Write an event log record - in debug mode we will also
        # see this message printed.
        try:
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
                )
        except error:
            pass #Log is most likely full, we do not want to die on this


        num_connections = 0
        #Waiting for an event
        while 1:
            pipeHandle = CreateNamedPipe(MY_PIPE_NAME,
                    PIPE_ACCESS_DUPLEX| FILE_FLAG_OVERLAPPED,
                    PIPE_TYPE_MESSAGE | PIPE_READMODE_BYTE,
                    PIPE_UNLIMITED_INSTANCES,       # max instances
                    0, 0, 6000,
                    self.CreatePipeSecurityObject())
            try:
                hr = ConnectNamedPipe(pipeHandle, self.overlapped)
            except error as details:
                print("Error connecting pipe!", details)
                CloseHandle(pipeHandle)
                break
            if hr==winerror.ERROR_PIPE_CONNECTED:
                # Client is already connected - signal event
                SetEvent(self.overlapped.hEvent)
            rc = WaitForMultipleObjects((self.hWaitStop, self.overlapped.hEvent), 0, INFINITE)
            if rc==WAIT_OBJECT_0:
                # Stop event, exit loop
                break
            else:
                # Pipe event - spawn thread to deal with it.
                thread.start_new_thread(self.ProcessClient, (pipeHandle,))
                num_connections = num_connections + 1

        # Sleep to ensure that any new threads are in the list, and then
        # wait for all current threads to finish.
        # What is a better way?
        Sleep(500)
        while self.thread_handles:
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING, 5000)
            print("Waiting for %d threads to finish..." % (len(self.thread_handles)))
            WaitForMultipleObjects(self.thread_handles, 1, 3000)
        # Write another event log record.
        try:
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STOPPED,
                (self._svc_name_, " after processing %d connections" % (num_connections,))
                )

        except error:
            pass #Log is most likely full, we do not want to die on this


if __name__=='__main__':
    win32serviceutil.HandleCommandLine(WPKGControlService)
