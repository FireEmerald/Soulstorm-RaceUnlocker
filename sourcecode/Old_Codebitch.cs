/*
 * ----------------------------------------------------------------------------
 * "THE BEER-WARE LICENSE" :
 * <Belial2003@gmail.com> wrote this file. As long as you retain this notice you
 * can do whatever you want with this stuff. If we meet some day, and you think
 * this stuff is worth it, you can buy me a beer in return. If you think we will not 
 * meet some day, you can also send me some. n0|Belial2003
 * ----------------------------------------------------------------------------
 */

using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using System.IO;

namespace SSRaceUnlocker
{
    class Codebitch
    {
        private string ClassicKey, WAKey, DCKey, SSKey, userRoot;
        //Set-Methods  START: used for saving entered Keys into private vars of this class
        public void SetClassicKey(String classicKey)
        {
            ClassicKey = classicKey;
        }
        public void SetWAKey(String waKey)
        {
            WAKey = waKey;
        }
        public void SetDCKey(String dcKey)
        {
            DCKey = dcKey;
        }
        public void SetuserRoot(String userroot)
        {
            userRoot = userroot;
        }
        //Set-Methods END

        private void Determine3264bitOS()
        {
            string SOFTWARE_KEY = "Software";
            string COMPANY_NAME = "Wow6432Node";
            RegistryKey win3264 = Registry.LocalMachine.OpenSubKey(SOFTWARE_KEY, false).OpenSubKey(COMPANY_NAME, false);
            if (win3264 != null)
            {
                SetuserRoot("HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\THQ\\");
            }
            else SetuserRoot("HKEY_LOCAL_MACHINE\\SOFTWARE\\THQ\\");
        }

        //Get-Methods START: used in RegWriter() to get desired values of vars and registry entries
        private String getClassicKey()
        {
            return ClassicKey;
        }
        private String getWAKey()
        {
            return WAKey;
        }
        private String getDCKey()
        {
            return DCKey;
        }
        private String getSSKey()
        {
            SSKey = (string)Registry.GetValue(getuserRoot() + "Dawn of War - Soulstorm", "CDKEY", RegistryValueKind.String); //TODO SSKey auslesen (32/64bit anpassung)
            return SSKey;
        }
        private String getuserRoot()
        {
            return userRoot;
        }
        private String getInstallLocation() 
        {
            String IL = (string)Registry.GetValue(getuserRoot() + "Dawn of War - Soulstorm", "InstallLocation", RegistryValueKind.String);

            return IL;
        }
        //Get-Methods END

        //RegWriter START: writes Registry entries and also creates the fake exe files (should have been sepereated, bad style, i know, and now STFU
        public void RegWriter()
        {
            Determine3264bitOS();
            
            string DoW = userRoot + "Dawn of War"; //registry path to DoW+WA
            string DoWDC = userRoot + "Dawn of War - Dark Crusade"; //registry path to DC
            string DoWSS = userRoot + "Dawn of War - Soulstorm"; //registry path to SS
            
            //Classic+WA Stuff START: Writes the keys of Classic and WA into the registry + creates the fake exe files
            if (getClassicKey() != null)
            {
                //Writes following registry entries into Classic registry dir: Classic CD-Key, InstallLocation of SS
                Registry.SetValue(DoW, "CDKEY", getClassicKey(), RegistryValueKind.String);
                Registry.SetValue(DoW, "InstallLocation", getInstallLocation() + "\\", RegistryValueKind.String); //the Backslash at the end of the path here is a MUST, verification won't work otherwise => Reric make bug ;)
                
                //Create fake Exe of DoW Classic START
                if (!File.Exists(getInstallLocation() + "\\W40k.exe"))
                {
                    File.Copy(getInstallLocation() + "\\GraphicsConfig.exe", getInstallLocation() + "\\W40k.exe"); //aint I a smart-ass? *g* fake exe like DCUnlock used, don't work, so this was the easiest solution i could come up with
                }
                //Create fake Exe of DoW Classic END

            }


            if (getWAKey() != null)
            {
                Registry.SetValue(DoW, "CDKEY_WXP", getWAKey(), RegistryValueKind.String);

                if (!File.Exists(getInstallLocation() + "\\W40kWA.exe"))
                {
                    File.Copy(getInstallLocation() + "\\GraphicsConfig.exe", getInstallLocation() + "\\W40kWA.exe");
                }
            }
            //Classic+WA Stuff END

            //DC Stuff START: Writes the key of DC into the registry + creates the fake exe file
            if (getDCKey() != null)
            {
                Registry.SetValue(DoWDC, "CDKEY", getDCKey(), RegistryValueKind.String);
                Registry.SetValue(DoWDC, "InstallLocation", getInstallLocation(), RegistryValueKind.String);

                if (!File.Exists(getInstallLocation() + "\\DarkCrusade.exe"))
                {
                    File.Copy(getInstallLocation() + "\\GraphicsConfig.exe", getInstallLocation() + "\\DarkCrusade.exe");
                }




            }
            //DC Stuff END

            //SS Stuff START: Writes the keys of Classic, WA and DC into the registry of Soulstorm
            if (getSSKey() != null)
            {

                    Registry.SetValue(DoWSS, "W40KCDKEY", getClassicKey(), RegistryValueKind.String);
                    Registry.SetValue(DoWSS, "WXPCDKEY", getWAKey(), RegistryValueKind.String);
                    Registry.SetValue(DoWSS, "DXP2CDKEY", getDCKey(), RegistryValueKind.String);
                 
            }
            //SS Stuff END
        }
        //RegWriter END


    }
}
