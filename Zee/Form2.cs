using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static iNovation.Code.General;
using static iNovation.Code.Desktop;
namespace Zee
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        private enum Profession 
        {
            Doctor, Pilot, Engineer, Writer, Hairdresser
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            TitleDrop(titleDropDown);
            GenderDrop(genderDropDown, false);
            BindProperty(professionDropDown, GetEnum(new Profession()), false);

        }

        private void resetButton_Click(object sender, EventArgs e)
        {
            Control[] controls = {titleDropDown, nameTextBox, genderDropDown, professionDropDown, yearsOfExperienceTextBox };            
            Clear(controls);
        }

        private void submitButton_Click(object sender, EventArgs e)
        {
            summaryTextBox.Text = Content(titleDropDown) + Environment.NewLine + Content(nameTextBox) + Environment.NewLine + Content(genderDropDown) + Environment.NewLine + Content(professionDropDown) + Environment.NewLine + Content(yearsOfExperienceTextBox);
        }

        private void copyToListBoxButton_Click(object sender, EventArgs e)
        {
            BindProperty(summaryListBox, StringToList(Content(summaryTextBox)));
        }

        private void upButton_Click(object sender, EventArgs e)
        {
            mFeedback("working", "this appears to be working");
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            ListsRemoveItem(summaryListBox);
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            ListsClearList(summaryListBox);
        }
    }
}
