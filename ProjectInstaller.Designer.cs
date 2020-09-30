namespace iParcer
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.serviceProcessInstalleriParser = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstalleriParser = new System.ServiceProcess.ServiceInstaller();
            // 
            // serviceProcessInstalleriParser
            // 
            this.serviceProcessInstalleriParser.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.serviceProcessInstalleriParser.Password = null;
            this.serviceProcessInstalleriParser.Username = null;
            // 
            // serviceInstalleriParser
            // 
            this.serviceInstalleriParser.Description = "Сервис обработки инвентаризационных данных";
            this.serviceInstalleriParser.DisplayName = "iParser";
            this.serviceInstalleriParser.ServiceName = "iParser";
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceProcessInstalleriParser,
            this.serviceInstalleriParser});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller serviceProcessInstalleriParser;
        private System.ServiceProcess.ServiceInstaller serviceInstalleriParser;
    }
}