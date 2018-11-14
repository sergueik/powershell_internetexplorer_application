#Copyright (c) 2018 Serguei Kouzmine
#
#Permission is hereby granted, free of charge, to any person obtaining a copy
#of this software and associated documentation files (the "Software"), to deal
#in the Software without restriction, including without limitation the rights
#to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#copies of the Software, and to permit persons to whom the Software is
#furnished to do so, subject to the following conditions:
#
#The above copyright notice and this permission notice shall be included in
#all copies or substantial portions of the Software.
#
#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
#THE SOFTWARE.


# https://archive.codeplex.com/?p=credentialmanagement
# based on: https://www.c-sharpcorner.com/forums/windows-credential-manager-with-c-sharp
# one can install wth nuget
# https://www.nuget.org/packages/CredentialManagement/
# https://stackoverflow.com/questions/14813370/how-to-access-the-stored-credentials-passwordvault-on-win7-and-win8

# NOTE: the example from MSDN https://blogs.msdn.microsoft.com/windowsappdev/2013/05/30/credential-locker-your-solution-for-handling-usernames-and-passwords-in-your-windows-store-app/
# is not possible to use with Windows 8 and Desktop apps - see
# the https://docs.microsoft.com/en-us/uwp/api/windows.security.credentials.passwordcredential classs is only defined for Windows 10
# cmdkey can create, list, and deletes stored user names and passwords or credentials.
# Passwords will not be displayed once they are stored
# https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/cmdkey
param(
  [switch]$debug
)

[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')

$shared_assemblies = @(
  'CredentialManagement.dll',
  'nunit.framework.dll'
)

$selenium_drivers_path = $shared_assemblies_path = "c:\Users\${env:USERNAME}\Downloads"

if (($env:SHARED_ASSEMBLIES_PATH -ne $null) -and ($env:SHARED_ASSEMBLIES_PATH -ne '')) {
  $shared_assemblies_path = $env:SHARED_ASSEMBLIES_PATH
}

pushd $shared_assemblies_path

$shared_assemblies | ForEach-Object { Unblock-File -Path $_; Add-Type -Path $_ }
popd

Add-Type -TypeDefinition @"
// "
using System;
using CredentialManagement;

public class Helper {
  private String password = null;
  private String userName = null;

  public string UserName {
    get { return userName; }
    set { userName = value; }
  }
  public string Password {
    set { password = value; }
  }

  public void SavePassword() {
    try {
      using (var cred = new Credential()) {
        cred.Password = password;
        cred.Target = userName;
        cred.Type = CredentialType.Generic;
        cred.PersistanceType = PersistanceType.LocalComputer;
        cred.Save();
      }
    } catch(Exception ex){
      Console.Error.WriteLine("Exception (ignord) " + ex.ToString());
    }
  }

  public String GetPassword() {
    try {
      using (var cred = new Credential()) {
        cred.Target = userName;
        cred.Load();
        return cred.Password;
      }
    } catch (Exception ex) {
      Console.Error.WriteLine("Exception (ignord) " + ex.ToString());
    }
    return "";
  }
}

"@  -ReferencedAssemblies 'System.Security.dll', "c:\Users\${env:USERNAME}\Downloads\CredentialManagement.dll"


$o = new-object Helper
$o.UserName = 'username'
$o.Password = 'newtest'
$o.SavePassword()
write-output ("Password loaded back: ")
$o.GetPassword()

<#
$o.Login()

Credential saved = new Credential("username", "password", "MyApp", CredentialType.Generic);
    saved.PersistanceType = PersistanceType.LocalComputer;
    saved.Save();

#>

# see cmdkey:
# https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-R2-and-2012/cc754243(v=ws.11)
