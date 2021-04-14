using System;
using System.Linq;
using NUnit.Framework;
using System.Threading;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using FlaUI.Core.Input;

namespace Tests
{
    public class Tests
    {
        private Application app;
        private Window window;

       


        private ConditionFactory cf;
       
        private readonly string appPath = @"C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.exe";

        private const string APP_TITLE = "Dokument1 - Microsoft Word u¿ytek niekomercyjny";

        private int sleepTimeShort = 500;
        private int sleepTimeNormal = 2000;
        private int sleepTimeLong = 8000;

        [SetUp]
        public void Setup()
        {
            app = FlaUI.Core.Application.Launch(appPath);
            app.WaitWhileMainHandleIsMissing(TimeSpan.FromMilliseconds(sleepTimeLong));
            using (var automation = new UIA3Automation())
            {
                cf = new ConditionFactory(new UIA3PropertyLibrary());
                window = app.GetAllTopLevelWindows(automation).First();
            }
        }

        [Test]
        public void UruchomienieProgramu()
        {
            Assert.AreEqual(window.Title, APP_TITLE, "Asercja 01: czy tytul okna jest poprawny");
        }

        [Test]
        public void OtwarcieOknaDialogowego()
        {
            var button = window.FindFirstDescendant(cf.ByName("Zapisz"));
            button.Click();
            Thread.Sleep(sleepTimeNormal);
            Window window2 = window.ModalWindows[0];
            Assert.AreEqual("Zapisz jako", window2.Title, "Asercja 02: czy tytul okna dialogowego jest poprawny");
        }
        [Test]
        public void Zapisaniepliku()
        {
            Keyboard.Type("Idzie kotek i skacze");
            OtwarcieOknaDialogowego();
            Window window2 = window.ModalWindows[0];
            var comboBoxes = window2.FindAllDescendants(cf.ByControlType(ControlType.ComboBox));
            TextBox wpisywanieNazwyPliku;
            foreach(var comboBox in comboBoxes)
            {
                if(comboBox.Name== "Nazwa pliku:")
                {
                    wpisywanieNazwyPliku = comboBox.FindFirstChild().As<TextBox>();
                    wpisywanieNazwyPliku.Text = "Kot";
                    break;
                }

            }

            var button = window.FindFirstDescendant(cf.ByName("Zapisz"));
            button.Click();

            Thread.Sleep(sleepTimeNormal);

           
           
        }
        [Test]
        public void OtwarcieZapisanegoPliku()
        {
            var button = window.FindFirstDescendant(cf.ByName("Przycisk pakietu Office"));
            button.Click();
            var buttonOpen= window.FindFirstDescendant(cf.ByName("Otwórz"));
            buttonOpen.Click();
            buttonOpen.Click();
            Thread.Sleep(sleepTimeNormal);
            Window window2 = window.ModalWindows[0];

            var comboBoxes = window2.FindAllDescendants(cf.ByControlType(ControlType.ComboBox));
            TextBox wpisywanieNazwyPliku;
            foreach (var comboBox in comboBoxes)
            {
                if (comboBox.Name == "Nazwa pliku:")
                {
                    wpisywanieNazwyPliku = comboBox.FindFirstChild().As<TextBox>();
                    wpisywanieNazwyPliku.Text = "Kot.docx";
                    break;
                }

            }
            var buttonOpen2 = window.FindAllDescendants(cf.ByName("Otwórz"))[2];
            buttonOpen2.Click();
            //using (var automation = new UIA3Automation()) ;
            //var checkText = window.FindFirstDescendant(cf.ByText("Idzie kotek i skacze")).AsTextBox().Text;
            //System.Console.WriteLine(checkText);

            Thread.Sleep(sleepTimeNormal);
           
            
        }



        [TearDown]
        public void Teardown()
        {
            Thread.Sleep(sleepTimeNormal);
            app.Close();

        }
    }
}