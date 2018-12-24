using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using Word = Microsoft.Office.Interop.Word;


namespace MyWpfWordApp
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        #region INotify Changed Properties  
        private string message;
        public string Message
        {
            get { return message; }
            set { SetField(ref message, value, nameof(Message)); }
        }

        private string fullAddressSender;
        public string FullAddressSender
        {
            get { return fullAddressSender; }
            set { SetField(ref fullAddressSender, value, nameof(FullAddressSender)); }
        }

        private string firstName;
        public string FirstName
        {
            get { return firstName; }
            set { SetField(ref firstName, value, nameof(FirstName)); }
        }

        private string surname;
        public string Surname
        {
            get { return surname; }
            set { SetField(ref surname, value, nameof(Surname)); }
        }

        private string streetHouseNumber;
        public string StreetHouseNumber
        {
            get { return streetHouseNumber; }
            set { SetField(ref streetHouseNumber, value, nameof(StreetHouseNumber)); }
        }

        private string postcodeCity;
        public string PostcodeCity
        {
            get { return postcodeCity; }
            set { SetField(ref postcodeCity, value, nameof(PostcodeCity)); }
        }

        // Template for a new INotify Changed Property
        // for using CTRL-R-R
        private bool xxx;
        public bool Xxx
        {
            get { return xxx; }
            set { SetField(ref xxx, value, nameof(Xxx)); }
        }
        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            Message = "Start";

            FullAddressSender = "IBM, Armonk, New York, U.S.";
            FirstName = "Microsoft";
            Surname = "Headquarters";
            StreetHouseNumber = "One Microsoft Way";
            PostcodeCity = "Redmond, WA 98052";
        }

        /******************************/
        /*       Button Events        */
        /******************************/
        #region Button Events

        /// <summary>
        /// Button_1_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_1_Click(object sender, RoutedEventArgs e)
        {
            if (CreateAndShowMyWordDocument())
            { 
                Message = "Document created";
                Console.Beep();
            }
        }
        private bool CreateAndShowMyWordDocument()
        {
            try
            {
                // Declaration of the necessary Word Interop variables
                Word.Application objWord = null;
                GetWordInstance(ref objWord);
                Word.Documents objDocs = objWord.Documents;
                Word.Document objDoc;

                // Declaration and initialization of other necessary variables
                object objFile = AppDomain.CurrentDomain.BaseDirectory + "MyWordTemplates.dotm";
                object objVisible = true;
                object objMissing = System.Reflection.Missing.Value;

                // Creation of the Word document
                objDoc = objDocs.Add(ref objFile, ref objMissing, ref objMissing, ref objVisible);
                object fname = String.Format("{0}{1}.docx",AppDomain.CurrentDomain.BaseDirectory + "MyWordDoc",Properties.Settings.Default.DocNum++);
                objDoc.SaveAs(ref fname, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);

                // Show the Word document
                objWord.Visible = true;
                objWord.ActiveWindow.Activate();

                // Fill the Word text placeholder with our bounded variables
                if (!SetWordBookmarks(objWord, objDoc.Bookmarks, "FullAddressSender", FullAddressSender)) return false;
                if (!SetWordBookmarks(objWord, objDoc.Bookmarks, "FirstName", FirstName)) return false;
                if (!SetWordBookmarks(objWord, objDoc.Bookmarks, "Surname", Surname)) return false;
                if (!SetWordBookmarks(objWord, objDoc.Bookmarks, "StreetHouseNumber", StreetHouseNumber)) return false;
                if (!SetWordBookmarks(objWord, objDoc.Bookmarks, "PostcodeCity", PostcodeCity)) return false;

                // Activate and save the document
                objDoc.Activate();
                objDoc.Save();

                // Release of the Interop resources
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objWord);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objDocs);

                return true;
            }
            catch (Exception ex)
            {
                Console.Beep();
                Console.Beep();
                Message = String.Format("EXCEPTION: {0}", ex.ToString());
                return false;
            }
        }
        public void GetWordInstance(ref Word.Application objWord)
        {
            objWord = null;
            try
            {
                // Get activ object
                objWord = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                if (objWord == null)
                {
                    try
                    {
                        // Create a new object
                        objWord = new Word.Application();
                    }
                    catch (Exception ex)
                    {
                        throw new System.Exception("Error in GetWordInstance: " + ex.ToString());
                    }
                }
            }
        }
        public bool SetWordBookmarks(Word.Application objWord, Word.Bookmarks objBookmarks,string bookMarkName,string bookMarkValue)
        {
            try
            {
                Word.Bookmark objBookmark = null;
                Word.Range objRange = null;
                object objName;

                // Name of Bookmark - Object in the Word Template
                objName = bookMarkName;

                // get the Bookmark
                objBookmark = objBookmarks.get_Item(ref objName);
                objRange = objBookmark.Range;

                // set new value of Bookmark and
                // write new Bookmark back to Word.Document
                objRange.Text = bookMarkValue;

                return true;
            }
            catch (Exception ex)
            {
                Console.Beep();
                Console.Beep();
                Message = String.Format("EXCEPTION: {0}", ex.ToString());
                return false;
            }
        }

        /// <summary>
        /// Button_Close_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion
        /******************************/
        /*      Menu Events          */
        /******************************/
        #region Menu Events

        #endregion
        /******************************/
        /*      Other Events          */
        /******************************/
        #region Other Events

        /// <summary>
        /// Lable_Message_MouseDown
        /// Clear Message
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Lable_Message_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Message = "";
        }

        /// <summary>
        /// Window_Closing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // We remember the last DocNum, for next program start
            Properties.Settings.Default.Save();
        }

        #endregion
        /******************************/
        /*      Other Functions       */
        /******************************/
        #region Other Functions

        /// <summary>
        /// SetField
        /// for INotify Changed Properties
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="field"></param>
        /// <param name="value"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        protected bool SetField<T>(ref T field, T value, string propertyName)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
        private void OnPropertyChanged(string p)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
        }

        #endregion
    }
}
