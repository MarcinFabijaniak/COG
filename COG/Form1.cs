using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Tekla.Structures.Model;
using TSM=Tekla.Structures.Model;

namespace COG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string fileToRead;
        DataTable dataTable = new DataTable("COG");

        private static void newObject(string fromModelCUP, double fromModelCOG_X, double fromModelCOG_Y, double fromModelCOG_Z)
        {
            TSM.Beam newObject = new TSM.Beam();

            newObject.Name = "COG for: " + fromModelCUP;
            newObject.Profile.ProfileString = "D100";
            newObject.Material.MaterialString = "Z35";
            newObject.Class = "1000";
            newObject.PartNumber.Prefix = "cg";
            newObject.AssemblyNumber.Prefix = "CG";
            newObject.StartPoint.X = fromModelCOG_X;
            newObject.StartPoint.Y = fromModelCOG_Y;
            newObject.StartPoint.Z = fromModelCOG_Z;
            newObject.EndPoint.X = fromModelCOG_X + 10000;
            newObject.EndPoint.Y = fromModelCOG_Y;
            newObject.EndPoint.Z = fromModelCOG_Z;
            newObject.Position.Rotation = Position.RotationEnum.TOP;
            newObject.Position.Plane = Position.PlaneEnum.MIDDLE;
            newObject.Position.Depth = Position.DepthEnum.MIDDLE;

            newObject.Insert();
        }        

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void openCOGFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Model MyModel = new Model();

            // Always remember to check that you really have working connection
            if (MyModel.GetConnectionStatus())
            {
                //opens file dialog and lets user select file to open then reds file name
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = "d:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                int size = -1;
                if (openFileDialog.ShowDialog() == DialogResult.OK) // Test result.
                {
                    try
                    {
                        fileToRead = openFileDialog.FileName;
                        string text = File.ReadAllText(fileToRead);
                        size = text.Length;
                        textBox1.Text = fileToRead;
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("Error: Could not read file from disk. Original error: ");
                    }
                }

                StreamReader reader = new StreamReader(@fileToRead);

                List<string> listA = new List<string>();
                List<string> listB = new List<string>();
                List<string> listC = new List<string>();
                List<string> listD = new List<string>();
                List<string> listE = new List<string>();
                List<string> listF = new List<string>();

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    listA.Add(values[0]);
                    listB.Add(values[1]);
                    listC.Add(values[2]);
                    listD.Add(values[3]);
                    listE.Add(values[4]);
                    listF.Add(values[5]);
                }

                DataTable dtLists = new DataTable("Lists");

                // Add columns to the DataTable.    
                dtLists.Columns.Add("GUID");
                dtLists.Columns.Add("Cast Unit Pos");
                dtLists.Columns.Add("Fabrication Code");
                dtLists.Columns.Add("COG X");
                dtLists.Columns.Add("COG Y");
                dtLists.Columns.Add("COG Z");

                for (int i = 0; i < listA.Count; i++)
                {
                    DataRow dataRow = dtLists.NewRow();
                    dataRow["GUID"] = listA[i];
                    dataRow["Cast Unit Pos"] = listB[i];
                    dataRow["Fabrication Code"] = listC[i];
                    dataRow["COG X"] = listD[i];
                    dataRow["COG Y"] = listE[i];
                    dataRow["COG Z"] = listF[i];

                    dtLists.Rows.Add(dataRow);
                }
                dataTable = dtLists;
                dataGridView1.DataSource = dataTable;

                int rowsCount = dataTable.Rows.Count;

                textBox2.Text = rowsCount + " " + "COGs will be created";
            }

            else
            {
                MessageBox.Show("Model not open or version incompatible!", "COG Message");
            }
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Model MyModel = new Model();
            TSM.Beam newObject = new TSM.Beam();

            int rowsCount = dataTable.Rows.Count;
            //string fabCode = "";

            if (rowsCount == 0)
            {
                MessageBox.Show("Open the file to proceed!", "COG Message");
            }
            else
            {
                progressBar1.Minimum = 1;
                progressBar1.Maximum = rowsCount;
                progressBar1.Value = 1;
                progressBar1.Step = 1;

                for (int i = 0; i < rowsCount; i++)
                {
                    newObject.Name = "COG for: " + dataTable.Rows[i]["Fabrication Code"].ToString();
                    //fabCode = newObject.Name;
                    //Console.WriteLine(newObject.Name + " "  + fabCode);
                    //newObject.SetUserProperty("FABRICATION_CODE", fabCode);
                    newObject.Profile.ProfileString = "ELD1*1*100*100";//ELD1*15*100*100 d100
                    newObject.Material.MaterialString = "Z35";
                    newObject.Class = "1001";
                    newObject.PartNumber.Prefix = "cog";
                    newObject.AssemblyNumber.Prefix = "COG";
                    newObject.StartPoint.X = Convert.ToDouble(dataTable.Rows[i]["COG X"]);
                    newObject.StartPoint.Y = Convert.ToDouble(dataTable.Rows[i]["COG Y"]);
                    newObject.StartPoint.Z = Convert.ToDouble(dataTable.Rows[i]["COG Z"]);
                    newObject.EndPoint.X = Convert.ToDouble(dataTable.Rows[i]["COG X"]);
                    newObject.EndPoint.Y = Convert.ToDouble(dataTable.Rows[i]["COG Y"]);
                    newObject.EndPoint.Z = Convert.ToDouble(dataTable.Rows[i]["COG Z"]) + 100;
                    newObject.Position.Rotation = Position.RotationEnum.TOP;
                    newObject.Position.Plane = Position.PlaneEnum.MIDDLE;
                    newObject.Position.Depth = Position.DepthEnum.MIDDLE;

                    newObject.Insert();
                    progressBar1.PerformStep();
                    textBox2.Text = rowsCount + " " + "COGs created";
                }
                MyModel.CommitChanges();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
    
