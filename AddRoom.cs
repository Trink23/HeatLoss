using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace HeatLoss
{
    
    public partial class AddRoom : Form
    {
        String heightText;
        String AreaText;
        ImportExcel importInfo;
        int tempDif = 0;
        String nameTab;
        int noWall = 0;
        int idWall = 0;



        public AddRoom()
        {
            InitializeComponent();
            
            importInfo = new ImportExcel();

            importInfo.ImportTab(roomInfo, "Rooms");

            partRoomListBox.DataSource = importInfo.ListSheetInExcel();






            //Setting text inside of textBox for heightBox
            HeightBox.Text = "Height";
            HeightBox.ForeColor = Color.Gray;

            //Setting text inside of textBox for AreaBox
            AreaBox.Text = "Area";
            AreaBox.ForeColor = Color.Gray;

            //To set Surface GridView

            nameTab = partRoomListBox.Items[partRoomListBox.SelectedIndex].ToString();

            importInfo.ImportTab(elementsGridView, nameTab);

            setTextBoxRoomElements(nameTab);

            //importInfo.SetUpDataGridView(elementsGridView);

            //initialize Sumamry DataGridView



        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (roomInfo.SelectedRows.Count > 0) // make sure user select at least 1 row 
            {
                string name = roomInfo.SelectedRows[0].Cells[0].Value + string.Empty;

                string temperature = roomInfo.SelectedRows[0].Cells[1].Value + string.Empty;

                List<string> airChange = new List<String>();

                for(int i = 2; i < roomInfo.ColumnCount; i++)
                {
                    airChange.Add(roomInfo.SelectedRows[0].Cells[i].Value + string.Empty);
                }

                CatListBox.DataSource = airChange;
                NameTextBox.Text = name;
                RoomTempBox.Text = temperature;
                
            }
        }

        private void areaRoomButton_Click(object sender, EventArgs e)
        {
            CommandsAutocad commands = new CommandsAutocad();
            this.Hide();

            double value = commands.GetPointsFromUserAndReturnArea();

            //value = Math.Round(value);

            this.Show();

            AreaBox.Text = System.Convert.ToString(value);

        }

        private void HeightBox_Enter(object sender, EventArgs e)
        {
            heightText = HeightBox.Text; // Stores old text value if you want
            HeightBox.Text = ""; // Clears the text field
            HeightBox.ReadOnly = false; // Makes the field editable

        }

        private void HeightBox_Leave(object sender, EventArgs e)
        {
            
            HeightBox.ForeColor = Color.Black;

        }

        

        private void AreaBox_Enter(object sender, EventArgs e)
        {
            AreaText = AreaBox.Text; // Stores old text value if you want
            AreaBox.Text = ""; // Clears the text field
            AreaBox.ReadOnly = false; // Makes the field editable

        }

        private void AreaBox_Leave(object sender, EventArgs e)
        {
            AreaBox.ForeColor = Color.Black;
        }

        /** 
         * Event to handle the first Watt calculation and added to the heatloss tab
         */

        private void AddRoom_Click(object sender, EventArgs e)
        {
            double volum = 0;            
            double watts = 0;
            float airChange = 0;
            tempDif = 0;
            
            int outsideTemp = 0;
            tabControl1.SelectTab(1);

            try
            {
                int roomTemp = int.Parse(RoomTempBox.Text.ToString());

                volum = float.Parse(AreaBox.Text.Replace(".", ",")) * float.Parse(HeightBox.Text.Replace(".", ","));

                if (hasChimneysBox.Checked)
                {
                    if (restrictorFittedBox.Checked)
                    {
                        if (volum < 70 & volum > 40)
                        {
                            airChange = 2;
                        }
                        else if (volum < 40)
                        {
                            airChange = 3;
                        }

                    }
                    else
                    {
                        if (volum < 70 & volum > 40)
                        {
                            airChange = 2;
                        }
                        else if (volum < 40)
                        {
                            airChange = 5;
                        }
                    }
                }
                else
                {

                    airChange = float.Parse(CatListBox.Items[CatListBox.SelectedIndex].ToString());

                }

                outsideTemp = int.Parse(OutsideTempBox.Text.ToString().Replace("-",""));                
                
             

                tempDif = roomTemp - outsideTemp;

            }
            catch(Exception es)
            {
                Console.Write(es.Data);
                MessageBox.Show("Please check input");
                return;
            }            

            watts = Calculations.WattsCalculate(volum,airChange,tempDif); ;

            wattsLabel.Text = watts + string.Empty;

            labelDimmension.Text = volum + string.Empty;

            textLengthArea.Text = AreaBox.Text;


        }

        private void partRoomListBox_Click(object sender, EventArgs e)
        {

            nameTab = partRoomListBox.Items[partRoomListBox.SelectedIndex].ToString();

            importInfo.ImportTab(elementsGridView, nameTab);

            setTextBoxRoomElements(nameTab);

        }
        private void setTextBoxRoomElements(String namePartRoom)
        {
            
            if(namePartRoom.Equals("Ceilings") || namePartRoom.Equals("Floors"))
            {
               lengthButton.Visible = false;

               textLengthArea.Text = AreaBox.Text;
            }
            else
            {

               textLengthArea.Text = "";

               lengthButton.Visible = true;
                
            }
        }

        private void elementsGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (elementsGridView.SelectedRows.Count > 0) // make sure user select at least 1 row 
            {
                float uValue = float.Parse(elementsGridView.SelectedRows[0].Cells[1].Value.ToString());

                txtElement.Text = elementsGridView.SelectedRows[0].Cells[0].Value.ToString();

                textUValue.Text = uValue + String.Empty;

                textCelsious.Text = tempDif + String.Empty;
            }

        }

        private void areaFloor_Click(object sender, EventArgs e)
        {
            CommandsAutocad commands = new CommandsAutocad();
            this.Hide();

            double value = commands.GetPointsFromUserAndReturnArea();

            //value = Math.Round(value);

            this.Show();

            textLengthArea.Text = System.Convert.ToString(value);

        }

       

        private void lengthButton_Click(object sender, EventArgs e)
        {
            this.Hide();

            double value = new CommandsAutocad().GetPointsFromUserAndReturnLine();

            this.Show();

            textLengthArea.Text = System.Convert.ToString(value);
        }

        private void AddRoom_FormClosing(object sender, FormClosingEventArgs e)
        {
            Heatloss.GetUserInterface().Show();
        }

        /** 
         * Event to add heatloss to the summary of the room 
         * also modify the watts accordingly
         */

        private void addHeatLossToSumaryButton_Click(object sender, EventArgs e)
        {
            float lengthArea;
            float uValue;
            int tempDifference;
            float height;
            String element;
            
            double Watts = 0;

            double total;
           

            try
            {
                element = txtElement.Text;
                lengthArea = float.Parse(textLengthArea.Text);
                tempDifference = int.Parse(textCelsious.Text);
                uValue = float.Parse(textUValue.Text);
                height = float.Parse(HeightBox.Text);

            }
            catch(Exception a)
            {
                Console.Write(a.ToString());
                MessageBox.Show("Please check input");
                return;
            }

            /*
             * Add an external wall to the sumary tab
             */
             
            if (nameTab.Equals("ExternalWalls"))
            {
                lengthArea = lengthArea * height;

                Watts = Calculations.WallWattsCalculate(lengthArea, tempDifference, uValue);
               
                sumaryDataGridView.Rows.Insert(noWall,idWall, "External Wall", element, lengthArea, uValue, Watts);

                wattsLabel.Text = float.Parse(wattsLabel.Text) + Watts + String.Empty;

                if(doorCheckBox.Checked)
                {
                    try
                    {
                        lengthArea = float.Parse(doorWidthBox.Text) * float.Parse(doorHeightBox.Text);
                    }
                    catch(Exception A)
                    {
                        MessageBox.Show("Please check input");
                        return;
                    }

                    noWall++;

                    Watts = lengthArea * 2.10;

                    sumaryDataGridView.Rows.Insert(noWall, "","External Wall", "Door", lengthArea, "2,10", Watts);

                    wattsLabel.Text = float.Parse(wattsLabel.Text) + Watts + String.Empty;

                    doorCheckBox.Checked = false;
                }

                if(glazingCheckBox.Checked)
                {
                    try
                    {
                        lengthArea = float.Parse(glazingWidthBox.Text) * float.Parse(glazingHeightBox.Text);
                    }
                    catch (Exception A)
                    {
                        MessageBox.Show("Please check input");
                        return;
                    }

                    noWall++;

                   

                    Watts = lengthArea * 2.10;

                    sumaryDataGridView.Rows.Insert(noWall, "", "External Wall", "Glazing", lengthArea, "2,10", Watts);

                    wattsLabel.Text = float.Parse(wattsLabel.Text) + Watts + String.Empty;

                    glazingCheckBox.Checked = false;

                }

            }

            /*
             * Add an internal wall to the sumary tab
             */

            if(nameTab.Equals("InternalWalls") || nameTab.Equals("PartyWalls"))
            {
                lengthArea = lengthArea * height;

                Watts = Calculations.WallWattsCalculate(lengthArea, tempDifference, uValue);

                sumaryDataGridView.Rows.Insert(noWall, idWall, "Internal Wall", element, lengthArea, uValue, Watts * 0);

                wattsLabel.Text = float.Parse(wattsLabel.Text) + (Watts * 0) + String.Empty;

            }

            /*
             * Add an Ceilings  to the sumary tab
             */

            if (nameTab.Equals("Ceilings"))
            {
                

                Watts = Calculations.WallWattsCalculate(lengthArea, tempDifference, uValue);

                sumaryDataGridView.Rows.Insert(noWall, idWall, "Ceiling", element, lengthArea, uValue, Watts);

                wattsLabel.Text = float.Parse(wattsLabel.Text) + Watts + String.Empty;

            }
            /*
            * Add an  Floors to the sumary tab
            */

            if (nameTab.Equals("Floors"))
            {


                Watts = Calculations.WallWattsCalculate(lengthArea, tempDifference, uValue);

                sumaryDataGridView.Rows.Insert(noWall, idWall, "Floors", element, lengthArea, uValue, Watts);

                wattsLabel.Text = float.Parse(wattsLabel.Text) + Watts + String.Empty;

            }


            noWall++;
            idWall++;
        }

        /**
         * If door CheckBox is clicked
         */

        private void doorCheckBox_Click(object sender, EventArgs e)
        {
            if(doorCheckBox.Checked)
            {
                doorGlazing.Visible = true;

                doorHeight.Visible = true;

                doorHeightBox.Visible = true;

                doorWidthBox.Visible = true;

                doorWidth.Visible = true;
                
            }
            else
            {
                doorHeight.Visible = false;

                doorHeightBox.Visible = false;

                doorWidthBox.Visible = false;

                doorWidth.Visible = false;
            }

            if(doorCheckBox.Checked == false & glazingCheckBox.Checked == false)
            {
                doorGlazing.Visible = false;
            }
        }

        /**
        * If glazing CheckBox is clicked
        */

        private void glazingCheckBox_Click(object sender, EventArgs e)
        {
            if (glazingCheckBox.Checked)
            {
                doorGlazing.Visible = true;

                glazingHeight.Visible = true;

                glazingWitdh.Visible = true;

                glazingHeightBox.Visible = true;

                glazingWidthBox.Visible = true;

            }
            else
            {
                glazingHeight.Visible = false;

                glazingWitdh.Visible = false;

                glazingHeightBox.Visible = false;

                glazingWidthBox.Visible = false;

            }

            if (doorCheckBox.Checked == false & glazingCheckBox.Checked == false)
            {
                doorGlazing.Visible = false;
            }

        }

        /**
         * Remove line from the sumary dataGridTab
         */

        private void removeButton_Click(object sender, EventArgs e)
        {
            int indexLine = 0;
            int idRemoved = -1;
            int maxNoRows = 0;
            int count = 0;

            if (sumaryDataGridView.SelectedRows.Count > 0) // make sure user select at least 1 row 
            {
                indexLine = sumaryDataGridView.SelectedRows[0].Index;
                maxNoRows = sumaryDataGridView.Rows.Count - 1;
                count = indexLine;
            }
            else
            {
                return;
            }

           

            while (idRemoved == -1)
            {
                
                if (sumaryDataGridView.Rows[count].Cells[0].Value.ToString().Equals(""))
                {
                    count--;
                }
                else
                {
                    idRemoved = int.Parse(sumaryDataGridView.Rows[count].Cells[0].Value.ToString());
                }
            }

            sumaryDataGridView.Rows.RemoveAt(indexLine);
            

            noWall = maxNoRows;
            idWall = idRemoved;
        }

        /**
         * Add room to General Sumary
         */
        private void doneButton_Click(object sender, EventArgs e)
        {

        }
    }
}
