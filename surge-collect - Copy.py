import wmi
import os
import sys
import getpass
import itertools
import time
import ctypes
import ConfigParser

#sets the globals and initializes console colors
STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE= -11
STD_ERROR_HANDLE = -12

FOREGROUND_PURPLE = 0x05 # text color contains purple.
FOREGROUND_AQUA = 0x03 # text color contains AQUA.
FOREGROUND_GREEN= 0x02 # text color contains green.
FOREGROUND_RED  = 0x04 # text color contains red.
FOREGROUND_WHITE = 0x0F # text color contains white.
FOREGROUND_INTENSITY = 0x08 # text color is intensified.

std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)

#gets the console color
def SetColor(color, handle=std_out_handle):

    bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)
    return bool

def main():
    config = ConfigParser.ConfigParser()
    config.read('surge.ini')
    hostname = raw_input('Enter host to collect from: ')
    username = raw_input('User name for host: ')
    pwd = getpass.getpass(prompt='User password: ')
    
    #check if surge-collect password is set
    p = config.get('DEFAULT', 'password')

    if p == 'False':
        p = getpass.getpass(prompt='surge-collect password: ')

    cmd = r'c:\Temp\Surge\surge-collect.exe %s' % p
    
    #set command-line options
    log = config.get('OPTIONS', 'log')
    mem = config.get('OPTIONS', 'no_mem')
    page = config.get('OPTIONS', 'pagefiles')
    file = config.get('OPTIONS', 'file')
    files = config.get('OPTIONS', 'files')

    if log != 'False':
        cmd = cmd + ' --log=%s' % log
    
    if mem == 'True':
        cmd = cmd + ' --no-memory'
    
    if page == 'True':
        cmd = cmd + ' --pagefiles'
    
    if file != 'False':
        cmd = cmd + ' --file=%s' % file
    
    if files != 'False':
        cmd = cmd + ' --file=%s' % files

    #file format and encryption
    storage = config.get('SECURITY', 'format')

    if storage.lower() == 'zip-pgp':
        pgp = config.get('SECURITY', 'recipient')
        cmd = cmd + ' --format=zip-pgp' + ' --recipient=%s' % pgp

    elif storage.lower() == 'zip':
        cmd = cmd + ' --format=zip'
    
    else:
        cmd = cmd + ' --format=%s' % storage
        
    #capture location settings
    directory = config.get('ARGUMENTS', 'directory')
    server = config.get('ARGUMENTS', 'host')
    s3 = config.get('ARGUMENTS', 's3')

    if directory != 'False':
        cmd = cmd + ' %s' % directory
    
    if server != 'False':
        insecure = config.get('SECURITY', 'insecure')
        cacert = config.get('SECURITY', 'cacert')
    
        if insecure == 'True':
            cmd = cmd + ' --insecure'
        else:
            cmd = cmd + r' --cacert=c:\Temp\Surge\cacert.pem'
        cmd = cmd + ' %s' % server
        
    if s3 != 'False':
        access = config.get('SECURITY', 's3_access')
        secret = config.get('SECURITY', 's3_secret')
        cmd = cmd + ' --s3-access-key=%s' % access
        cmd = cmd + ' --s3-secret-key=%s' % secret
        cmd = cmd + ' %s' % s3
    
    c = wmi.WMI()

    try:
        b = wmi.WMI(hostname, user=username, password=pwd)
    except Exception as e:
        if 'The RPC server is unavailable. ' in str(e):
            SetColor(FOREGROUND_RED | FOREGROUND_INTENSITY)
            print '\nThe RPC server is unavailable.'
            SetColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
            print '[done]'
            SetColor(FOREGROUND_WHITE)
        elif 'Access is denied. ' in str(e):
            SetColor(FOREGROUND_RED | FOREGROUND_INTENSITY)
            print '\nAccess is denied.'
            SetColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
            print '[done]'
            SetColor(FOREGROUND_WHITE)
        sys.exit(0)
    
#    dir_path = os.path.dirname(os.path.realpath(__file__))+'\\'+os.path.dirname(sys.argv[0])
    dir_path = os.path.dirname(os.path.realpath(__file__))
#    print(dir_path)

    SetColor(FOREGROUND_AQUA | FOREGROUND_INTENSITY)
    print "\nCopying surge-collect to %s" % hostname
    SetColor(FOREGROUND_WHITE)
    c.Win32_Process.Create(CommandLine=r'robocopy %s \\%s\c$\Temp\Surge surge-collect.exe /LOG+:%s\robocopy.log /R:0' % (dir_path,hostname,dir_path))
    while True:
        n = c.Win32_Process (Name="robocopy.exe")
        if not n:
            SetColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
            print'[done]\n'
            SetColor(FOREGROUND_WHITE)
            break

    if insecure == 'False':
        SetColor(FOREGROUND_AQUA | FOREGROUND_INTENSITY)
        print "\nCopying cacert to %s" % hostname
        SetColor(FOREGROUND_WHITE)
        cacert_path = cacert.split('\\')
        cacert_path = '\\'.join(cacert_path[:len(cacert_path)-1])
        c.Win32_Process.Create(CommandLine=r'robocopy %s \\%s\c$\Temp\Surge cacert.pem /LOG+:%s\robocopy.log /R:0' % (cacert_path,hostname,dir_path))
        while True:
            n = c.Win32_Process (Name="robocopy.exe")
            if not n:
                SetColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
                print'[done]\n'
                SetColor(FOREGROUND_WHITE)
                break
          
    b.Win32_Process.Create(CommandLine=cmd)
    SetColor(FOREGROUND_AQUA | FOREGROUND_INTENSITY)
    sys.stdout.write('Collecting memory sample from %s...' % hostname)
    spinner = itertools.cycle(['|', '/', '-', '\\'])
    count = 0
    while True:
        try:
            sys.stdout.write(spinner.next())
            sys.stdout.flush()
            time.sleep(0.1)
            sys.stdout.write('\b')
            p = b.Win32_Process (Name="surge-collect.exe")
            count += 1
            if not p:
                if count < 10:
                    SetColor(FOREGROUND_RED | FOREGROUND_INTENSITY)
                    print '\nBad surge-collect password'
                SetColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
                print'\n[done]\n'
                SetColor(FOREGROUND_AQUA | FOREGROUND_INTENSITY)
                print'Removing surge-collect from %s' % hostname
                c.Win32_Process.Create(CommandLine=r'cmd /c rmdir /s /q \\%s\c$\Temp\Surge' % hostname)
                SetColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
                print'\n[done]\n'
                SetColor(FOREGROUND_WHITE)
                sys.exit(0)
        except KeyboardInterrupt:
            SetColor(FOREGROUND_RED | FOREGROUND_INTENSITY)
            print'\nKeyboardInterrupt. Cleanup in progress. Surge-collect will continue to collect memory sample.'
            SetColor(FOREGROUND_AQUA | FOREGROUND_INTENSITY)
            print'Killing surge-collect process on %s' % hostname
            b.Win32_Process.Create(CommandLine=r'cmd /c taskkill /f /im surge-collect.exe')
            time.sleep(2)
            SetColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
            print'[done]\n'
            SetColor(FOREGROUND_AQUA | FOREGROUND_INTENSITY)
            print'Removing surge-collect from %s' % hostname
            c.Win32_Process.Create(CommandLine=r'cmd /c rmdir /s /q \\%s\c$\Temp\Surge' % hostname)
            SetColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
            print'[done]\n'
            SetColor(FOREGROUND_WHITE)
            sys.exit(0)
        
if __name__ == "__main__":
    
    ctypes.windll.kernel32.SetConsoleTitleA("Surge-Collect")
    SetColor(FOREGROUND_PURPLE | FOREGROUND_INTENSITY)
    print ""
    print " @@@@@@   @@@  @@@  @@@@@@@    @@@@@@@@  @@@@@@@@              @@@@@@@   @@@@@@   @@@       @@@       @@@@@@@@   @@@@@@@  @@@@@@@  "
    print "@@@@@@@   @@@  @@@  @@@@@@@@  @@@@@@@@@  @@@@@@@@             @@@@@@@@  @@@@@@@@  @@@       @@@       @@@@@@@@  @@@@@@@@  @@@@@@@  "
    print "!@@       @@!  @@@  @@!  @@@  !@@        @@!                  !@@       @@!  @@@  @@!       @@!       @@!       !@@         @@!    "
    print "!@!       !@!  @!@  !@!  @!@  !@!        !@!                  !@!       !@!  @!@  !@!       !@!       !@!       !@!         !@!    "
    print "!!@@!!    @!@  !@!  @!@!!@!   !@! @!@!@  @!!!:!    @!@!@!@!@  !@!       @!@  !@!  @!!       @!!       @!!!:!    !@!         @!!    "
    print " !!@!!!   !@!  !!!  !!@!@!    !!! !!@!!  !!!!!:    !!!@!@!!!  !!!       !@!  !!!  !!!       !!!       !!!!!:    !!!         !!!    "
    print "     !:!  !!:  !!!  !!: :!!   :!!   !!:  !!:                  :!!       !!:  !!!  !!:       !!:       !!:       :!!         !!:    "
    print "    !:!   :!:  !:!  :!:  !:!  :!:   !::  :!:                  :!:       :!:  !:!   :!:       :!:      :!:       :!:         :!:    "
    print ":::: ::   ::::: ::  ::   :::   ::: ::::   :: ::::              ::: :::  ::::: ::   :: ::::   :: ::::   :: ::::   ::: :::     ::    "
    print ":: : :     : :  :    :   : :   :: :: :   : :: ::               :: :: :   : :  :   : :: : :  : :: : :  : :: ::    :: :: :     :    "
    print ""
    SetColor(FOREGROUND_WHITE)
    
    if ctypes.windll.shell32.IsUserAnAdmin():
            main()
    else:
        SetColor(FOREGROUND_RED | FOREGROUND_INTENSITY)
        print ""
        print "[-] needs to be run as admin" + "\n"
        SetColor(FOREGROUND_WHITE)
        time.sleep(3)
        os._exit(0)