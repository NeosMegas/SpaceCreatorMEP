#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SpaceCreatorMEP.Properties;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Resources;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

#endregion

namespace SpaceCreatorMEP
{
    public class App : IExternalApplication
    {
        AddInId addInId = new AddInId(new Guid("8f1e12f8-f8e8-4ec2-a611-de764ebbf55e"));

        public Result OnStartup(UIControlledApplication application)
        {
            RibbonPanel ribbonPanel = AutomationPanel(application);
            return Result.Succeeded;
        }

        public RibbonPanel AutomationPanel(UIControlledApplication application, string tabName = "")
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            RibbonPanel ribbonPanel = null;
            if (string.IsNullOrEmpty(tabName))
                ribbonPanel = application.CreateRibbonPanel("MEP Spaces");
            else
                ribbonPanel = application.CreateRibbonPanel(tabName, "MEP Spaces");
            AddButton(ribbonPanel, "Создать\nпространства", assemblyPath, "SpaceCreatorMEP.Command", "Создание пространств на основе помещений из связи");
            return ribbonPanel;
        }

        private void AddButton(RibbonPanel ribbonPanel, string buttonName, string path, string linkToCommand, string toolTip)
        {
            PushButtonData buttonData = new PushButtonData(
               buttonName,
               buttonName,
               path,
               linkToCommand);
            ContextualHelp contextualHelp = new ContextualHelp(ContextualHelpType.Url, Properties.Settings.Default.helpURL);
            buttonData.SetContextualHelp(contextualHelp);
            PushButton Button = ribbonPanel.AddItem(buttonData) as PushButton;
            Button.ToolTip = toolTip;
            Button.LargeImage = new BitmapImage(new Uri(@"pack://application:,,,/SpaceCreatorMEP;component/Spaces32.png", UriKind.RelativeOrAbsolute));
            //BitmapImage b = new BitmapImage();
            //b.BeginInit();
            //Resources.Spaces32
            //b.UriSource = new Uri("pack://application:,,,/SpaceCreatorMEP;component/Resources/Spaces32.png");
            //b.EndInit();
            //Button.LargeImage = b;
        }


        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
