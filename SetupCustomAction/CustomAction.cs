using System;
using System.Collections;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace SetupCustomAction
{
    /// <summary>
    /// インストーラ用のカスタムアクション
    /// </summary>
    [System.ComponentModel.RunInstaller(true)]
    public sealed class CustomAction : Installer
    {
        /// <summary>
        /// インストール時のカスタムアクション
        /// </summary>
        /// <param name="savedState">savedState</param>
        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            // Reset Outlook resiliency keys to ensure add-in is enabled after install
            ResetResiliencyKeys();

            //msiexec /i "OkanSetup.msi" SILENT=TRUE ALLUSERS=1 /quiet /norestart
            //ALLUSERS=1 で、すべてのユーザを対象にインストール
            if (Context.Parameters["silent"] == "TRUE") return;

            var outlookProcess = Process.GetProcessesByName("OUTLOOK");
            if (outlookProcess.Length <= 0) return;

            _ = MessageBox.Show("Outlookが起動しています。Outlookを終了してからインストールしてください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.ServiceNotification);
            throw new InstallException();
        }

        /// <summary>
        /// アンインストール時のカスタムアクション
        /// </summary>
        /// <param name="savedState">savedState</param>
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            try
            {
                //msiexec /x "OkanSetup.msi" DELCONF=TRUE /quiet /norestart
                if (Context.Parameters["delconf"] == "TRUE")
                {
                    try
                    {
                        var directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookOkan\\");
                        Directory.Delete(directoryPath, true);
                    }
                    catch (Exception)
                    {
                        //Do Nothing.
                    }

                    return;
                }

                //msiexec /x "OkanSetup.msi" SILENT=TRUE /quiet /norestart
                if (Context.Parameters["silent"] == "TRUE") return;

                var result = MessageBox.Show("設定を削除しますか？", "設定削除の確認", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.ServiceNotification);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        var directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookOkan\\");
                        Directory.Delete(directoryPath, true);
                    }
                    catch (Exception)
                    {
                        //Do Nothing.
                    }
                }
            }
            catch (Exception)
            {
                //Do Nothing.
            }
        }

        public override void Commit(IDictionary savedState)
        {
        }

        public override void Rollback(IDictionary savedState)
        {
        }

        /// <summary>
        /// Reset Outlook resiliency registry keys to ensure add-in is not disabled.
        /// Called during installation to clear any previous crash/disable history.
        /// </summary>
        private static void ResetResiliencyKeys()
        {
            try
            {
                var addinProgId = "OutlookOkan";
                var officeVersions = new[] { "16.0", "15.0" };

                foreach (var version in officeVersions)
                {
                    // 1. Register in DoNotDisableAddinList
                    var doNotDisablePath = $@"Software\Microsoft\Office\{version}\Outlook\Resiliency\DoNotDisableAddinList";
                    using (var key = Registry.CurrentUser.CreateSubKey(doNotDisablePath))
                    {
                        key?.SetValue(addinProgId, 1, RegistryValueKind.DWord);
                    }

                    // 2. Delete CrashingAddinList key entirely (clean slate)
                    try
                    {
                        Registry.CurrentUser.DeleteSubKey(
                            $@"Software\Microsoft\Office\{version}\Outlook\Resiliency\CrashingAddinList",
                            throwOnMissingSubKey: false);
                    }
                    catch (Exception) { }

                    // 3. Delete DisabledItems key entirely (clean slate)
                    try
                    {
                        Registry.CurrentUser.DeleteSubKey(
                            $@"Software\Microsoft\Office\{version}\Outlook\Resiliency\DisabledItems",
                            throwOnMissingSubKey: false);
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception)
            {
                // Silently fail - best effort during install
            }
        }
    }
}