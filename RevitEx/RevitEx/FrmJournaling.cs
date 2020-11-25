using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitEx
{
    public partial class FrmJournaling : Form
    {
        const double Precision = 0.00001;
        Journaling m_dataBuffer;

        public FrmJournaling(Journaling dataBuffer)
        {
            InitializeComponent();

            m_dataBuffer = dataBuffer;

            cmbType.DataSource = m_dataBuffer.WallTypes;
            cmbType.DisplayMember = "Name";
            cmbLevel.DataSource = m_dataBuffer.Levels;
            cmbLevel.DisplayMember = "Name";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Autodesk.Revit.DB.XYZ startPoint = new Autodesk.Revit.DB.XYZ(
                    Convert.ToDouble(txtX1.Text.Trim()),
                    Convert.ToDouble(txtY1.Text.Trim()),
                    Convert.ToDouble(txtZ1.Text.Trim()));

            Autodesk.Revit.DB.XYZ endPoint = new Autodesk.Revit.DB.XYZ(
                    Convert.ToDouble(txtX2.Text.Trim()),
                    Convert.ToDouble(txtY2.Text.Trim()),
                    Convert.ToDouble(txtZ2.Text.Trim()));

            if (startPoint.Equals(endPoint))
            {
                Autodesk.Revit.UI.TaskDialog.Show("Revit", "Start point should not equal end point.");
                return;
            }

            double diff = Math.Abs(startPoint.Z - endPoint.Z);
            if (diff > Precision)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Revit", "Z coordinate of start and end points should be equal.");
                return;
            }

            Autodesk.Revit.DB.Level level = cmbLevel.SelectedItem as Autodesk.Revit.DB.Level;  // level information
            if (null == level)  // assert it isn't null
            {
                Autodesk.Revit.UI.TaskDialog.Show("Revit", "The selected level is null or incorrect.");
                return;
            }

            Autodesk.Revit.DB.WallType type = cmbType.SelectedItem as Autodesk.Revit.DB.WallType;  // wall type
            if (null == type)    // assert it isn't null
            {
                Autodesk.Revit.UI.TaskDialog.Show("Revit", "The selected wall type is null or incorrect.");
                return;
            }

            // Invoke SetNecessaryData method to set the collected support data 
            m_dataBuffer.SetNecessaryData(startPoint, endPoint, level, type);

            // Set result information and close the form
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
