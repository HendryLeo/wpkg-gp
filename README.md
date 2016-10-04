# wpkg-gp mod
This is a modification for the original wpkg-gp intended to be used with [WPKG-GP Client](https://github.com/sonicnkt/wpkg-gp-client).

Added features not available in the original:
- query for pending tasks (```wpkgpipeclient.exe Query```)
- execute wpkg sychronisation without reboot (```wpkgpipeclient.exe ExecuteNoReboot```)
- blacklist systems from executing wpkg-gp:
  - add blacklist.txt to your wpkg root directory and add the name of the system (per line) that should be blocked.
  - lines starting with "#" will be ignored.
  - ```!all!``` will block all systems.

Running WPKG as a Group Policy Extension with a few modification to work with my other project (WPKG-GP Client) 

See original [project WIKI](https://github.com/cleitet/wpkg-gp/wiki) for compiling Instructions or download a precompiled release from this repository.
