using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace UpdateWordDoc
{
    public partial class Form1 : Form
    {
        public object Application { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void findAndReplace(Microsoft.Office.Interop.Word.Application app, Document doc, string findText, string replaceText)
        {
            //create the findObject
            Find findObject = app.Selection.Find;
            
            //ClearFormatting needs to be called on findObject && findObject.Replacement
            findObject.ClearFormatting();
            findObject.Text = findText;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replaceText;

            //define missingVal (object representation of null value)
            object missingVal = System.Reflection.Missing.Value;

            //defind replaceVal -> replaceAll
            object replaceVal = WdReplace.wdReplaceAll;
            findObject.Execute(ref missingVal, ref missingVal, ref missingVal, ref missingVal, ref missingVal, ref missingVal, ref missingVal, ref missingVal, ref missingVal, ref missingVal, ref replaceVal, ref missingVal, ref missingVal, ref missingVal, ref missingVal);
            
            //iterate all shapes in doc (SHAPES INCLUDE text objects, drawings, shapes, pictures, OLE objects, ActiveX controls, and callouts. EXCLUDES headers & footers) 
            Shapes shapes = doc.Shapes;
            foreach (Shape shape in shapes)
            {
                //if shape has text, find and replace
                if (shape.TextFrame.HasText != 0)
                {
                    string initialText = shape.TextFrame.TextRange.Text;
                    string resultText = initialText.Replace(findText, replaceText);
                    if (initialText != resultText)
                    {
                        shape.TextFrame.TextRange.Text = resultText;
                    }
                }
            }
        }

        private string getDocPath() {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.ShowDialog();
            return fileDialog.FileName.ToString(); 
        }

        private string copyAndRenameDoc(string position, string companyName) {
            //select the doc to open
            string docPath = getDocPath();

            //split off the file extension 
            string[] splitPathArr = docPath.Split('.');

            //build the new doc path
            string newDocPath = splitPathArr[0] + " - " + position + " - " + companyName + ".docx";

            //copy and rename the doc
            System.IO.File.Copy(docPath, newDocPath);

            //return the new path
            return newDocPath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //these are defined in the word doc for quick find and replace operations
            const string dateText = "{DATE}";
            const string positionUppercaseText = "{POSITION_NAME_UPPERCASE}";
            const string positionLowercaseText = "{POSITION_NAME_LOWERCASE}";
            const string hiringNameText = "{HIRING_NAME}";
            const string companyNameText = "{COMPANY_NAME}";
            const string addressText = "{ADDRESS}";
            const string companyActivityText = "{COMPANY_ACTIVITY}";
            const string skillsText = "{SKILLS}";

            //get data from text boxes
            string dateData = DateTime.Now.ToString("MM/dd/yyyy");
            string positionUppercaseData = positionBox.Text.ToUpper();
            string positionLowercaseData = positionBox.Text;
            string hiringNameData = nameBox.Text;
            string companyNameData = companyNameBox.Text;
            string addressData = addressBox.Text;
            string companyActivityData = companyActivityBox.Text;
            string skillsData = skillsBox.Text;

            //copy and rename the doc
            string newDocPath = copyAndRenameDoc(positionLowercaseData, companyNameData);

            //create Word application object
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc = app.Documents.Open(newDocPath);

            //find and replace the text in the new Word doc
            findAndReplace(app, doc, dateText, dateData);
            findAndReplace(app, doc, positionUppercaseText, positionUppercaseData);
            findAndReplace(app, doc, positionLowercaseText, positionLowercaseData);
            findAndReplace(app, doc, hiringNameText, hiringNameData);
            findAndReplace(app, doc, companyNameText, companyNameData);
            findAndReplace(app, doc, addressText, addressData);
            findAndReplace(app, doc, companyActivityText, companyActivityData);
            findAndReplace(app, doc, skillsText, skillsData);

            //save new doc, kill the process
            doc.Save();
            app.Quit();
        }
    }
}
