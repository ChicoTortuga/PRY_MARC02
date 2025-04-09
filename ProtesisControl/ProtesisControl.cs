using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using OfficeOpenXml;

namespace ProtesisControl
{
    public partial class ProtesisControl : Form
    {
        System.IO.Ports.SerialPort puerto;
        String[] listado_puerto = System.IO.Ports.SerialPort.GetPortNames();
        string datos_puerto;
        string posicion;
        float serie1 = 0, serie2 = 0;
        int tiempo = 0;
        int x = 0;
        int y = 0;
        string Xg;
        string Yg;
        bool IsOpen = false;
        System.Drawing.Rectangle r1 = new System.Drawing.Rectangle();
        System.Drawing.Point? prevPosition = null;
        ToolTip tooltip = new ToolTip();

        int valorPosicionRecibida;

        public ProtesisControl()
        {
            InitializeComponent();

            if (listado_puerto.Length <= 0)
            {
                comboBox1.Items.Add("NO PUERTO");
            }
            foreach (var item in listado_puerto)
            {
                comboBox1.Items.Add(item);
            }

            comboBox1.SelectedIndex = 0;
            posicion = "";

        }

        public void serial()
        {

            try
            {
                this.puerto = new System.IO.Ports.SerialPort("" + comboBox1.SelectedItem, 115200, System.IO.Ports.Parity.None, 8, System.IO.Ports.StopBits.One);
                this.puerto.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(recepcion);

            }
            catch (Exception)
            {
                MessageBox.Show("Verifique:" + System.Environment.NewLine + "- Voltage" + System.Environment.NewLine + "- Conexion del puerto", "Error de puerto COM", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        public void recepcion(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            try
            {
                //Thread.Sleep(10);
                datos_puerto = this.puerto.ReadLine();
                Console.WriteLine(datos_puerto);
                if (datos_puerto.StartsWith("$"))
                { this.Invoke(new EventHandler(actualizar)); }

            }

            catch (Exception) { }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            tiempo++;
            chart1.Series[0].Points.AddXY((tiempo), serie1);
            chart1.Series[1].Points.AddXY((tiempo), serie2);
            // chart1.Series[0].Points.AddXY(serie2, serie1); //ESTA ES LA FORMA DE GRAFICAR XY CON VALORES DIFERENTES
            label4.Text = "" + serie1;
            valorPosicionRecibida = Convert.ToInt32(serie2);
            switch (valorPosicionRecibida)
            {
                case 0:
                    posicion = "ABIERTA";
                    break;
                case 250:
                    posicion = "PINZA";
                    break;
                case 500:
                    posicion = "PUÑO";
                    break;
                default:
                    posicion = "DESCONOCIDO"; // Caso por defecto para valores no esperados
                    break;
            }
            label5.Text = posicion;

            listBox1.Items.Add(serie1 + ";" + serie2 + ";" + (tiempo)); ;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem.ToString() != "NO PUERTO")
            {
                serial();
                puerto.Open();
                pictureBox1.Image = Properties.Resources.tapones_de_enchufes_03;
                IsOpen = true;
                timer1.Start();
                listBox1.Items.Clear();
            }
            else
            {
                MessageBox.Show("NO EXISTE PUERTO");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (IsOpen)
            {
                puerto.Close();
                Console.WriteLine("CLOSED PORT");
                IsOpen = false;
                pictureBox1.Image = Properties.Resources.tapones_de_enchufes_02;
                timer1.Stop();
            }
        }

        private void chart1_PostPaint(object sender, System.Windows.Forms.DataVisualization.Charting.ChartPaintEventArgs e)
        {
            System.Drawing.Font drawFont = new System.Drawing.Font("Verdana", 8);
            System.Drawing.SolidBrush drawBrush = new System.Drawing.SolidBrush(System.Drawing.Color.Red);

            System.Drawing.StringFormat drawFormat = new System.Drawing.StringFormat();
            if (!(string.IsNullOrEmpty(Xg) && string.IsNullOrEmpty(Yg)))
            {
                e.ChartGraphics.Graphics.DrawString("ELong=" + Xg + "\n" + "Lectura=" + Yg, drawFont, drawBrush, x + 5, y - 2);
                e.ChartGraphics.Graphics.DrawRectangle(new Pen(Color.Red, 3), r1);
            }
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            var pos1 = e.Location;
            if (prevPosition.HasValue && pos1 == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos1;
            var results1 = chart1.HitTest(pos1.X, pos1.Y, false, ChartElementType.DataPoint);
            foreach (var result1 in results1)
            {
                if (result1.ChartElementType == ChartElementType.DataPoint)
                {
                    var prop1 = result1.Object as DataPoint;
                    if (prop1 != null)
                    {
                        var pointXPixel = result1.ChartArea.AxisX.ValueToPixelPosition(prop1.XValue);
                        var pointYPixel = result1.ChartArea.AxisY.ValueToPixelPosition(prop1.YValues[0]);

                        // check if the cursor is really close to the point (2 pixels around)

                        tooltip.Show("Tiempo=" + prop1.XValue + ", Lectura=" + prop1.YValues[0], this.chart1,
                                        pos1.X, pos1.Y - 15);

                    }
                }
            }
        }
        
        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            chart1.Invalidate();
            r1.X = x;
            r1.Y = y;
            r1.Width = 3;
            r1.Height = 3;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.FileName = "Umbrales_Lectura";
            dlg.Filter = "Archivo Excel (*.xlsx)|*.xlsx";
            DialogResult dlgResult = dlg.ShowDialog();
            if (dlgResult == DialogResult.OK)
            {
                ExcelPackage.License.SetNonCommercialPersonal("Pedro");
                // Crear un archivo Excel
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    // Crear una hoja de trabajo
                    var worksheet = package.Workbook.Worksheets.Add("Datos");

                    // Encabezados
                    worksheet.Cells[1, 1].Value = "Lectura";
                    worksheet.Cells[1, 2].Value = "Posicion";
                    worksheet.Cells[1, 3].Value = "Tiempo";

                    // Llenar datos
                    for (int j = 0; j < listBox1.Items.Count; j++)
                    {
                        string[] arreglo = listBox1.Items[j].ToString().Split(';');
                        worksheet.Cells[j + 2, 1].Value = Convert.ToDouble(arreglo[0]); // Columna 1
                        worksheet.Cells[j + 2, 2].Value = Convert.ToDouble(arreglo[1]); // Columna 2
                        worksheet.Cells[j + 2, 3].Value = Convert.ToInt32(arreglo[2]); // Columna 3
                    }

                    // Guardar el archivo
                    FileInfo fi = new FileInfo(dlg.FileName);
                    package.SaveAs(fi);

                    MessageBox.Show("Datos guardados en Excel", "Excel");
                }
            }
        }

        private void ProtesisControl_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (IsOpen)
            {
                puerto.Close();
                Console.WriteLine("CLOSED PORT");
                IsOpen = false;
                pictureBox1.Image = Properties.Resources.tapones_de_enchufes_02;
                timer1.Stop();
            }
        }

        public void actualizar(object s, EventArgs e)
        {
            try
            {
                datos_puerto = datos_puerto.Remove(0, 1);
                string[] arreglo = datos_puerto.Split(';');
                serie1 = (float)Convert.ToDouble(arreglo[0]);
                serie2 = (float)Convert.ToDouble(arreglo[1]);

            }

            catch (Exception) { }

        }
    }
}