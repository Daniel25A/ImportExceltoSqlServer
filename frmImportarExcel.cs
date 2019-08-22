using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Punto_de_Venta.Formularios
{
    public partial class frmImportarExcel : Form
    {
        static string Ruta = "";
        public frmImportarExcel()
        {
            InitializeComponent();
        }

        private void gimportacion_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void frmImportarExcel_Load(object sender, EventArgs e)
        {
            buscadorarchivo.Filter = "Archivos Excel(*.xlsx)|*.xlsx";
            lblimportando.Visible = false;
            this.Text = "IMPORTAR PRODUCTOS DE EXCEL";

        }
        void CargarColumnas(ComboBox Lista)
        {
            try
            {
                using (var Columns = new Clases.Consultas.CProductos())
                {
                    if (Columns.GetColumnas(Ruta, cmbhojas.Text.Trim()) == null) return;
                    else
                        Columns.GetColumnas(Ruta, cmbhojas.Text.Trim()).ForEach(x => Lista.Items.Add(x));
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocurrio un Error en " + ex.TargetSite);
            }
        }
        private void btnseleccionar_Click(object sender, EventArgs e)
        {
            cmbhojas.Text = string.Empty;
            gimportacion.Controls.OfType<ComboBox>().Where(x => x.Items.Count > 0).ToList<ComboBox>().ForEach(clear => { clear.Items.Clear(); clear.Text = string.Empty; });

            if (buscadorarchivo.ShowDialog() == DialogResult.OK)
            {
                Ruta = buscadorarchivo.FileName;
                using (Clases.Consultas.CProductos Hojas = new Clases.Consultas.CProductos())
                {
                    cmbhojas.Items.Clear();
                    if (Hojas.GetHojas(Ruta) == null)
                        return;
                    else
                        Hojas.GetHojas(Ruta).ForEach(x => cmbhojas.Items.Add(x));
                }
            }
        }

        private void btncargar_Click(object sender, EventArgs e)
        {
            if (cmbhojas.Items.Count == 0) return;
            if (cmbhojas.Text.Trim() == string.Empty) {
                MessageBox.Show("Por Favor, Seleccione la Hoja del Archivo Excel a Importar", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            foreach (ComboBox comBoXX in gimportacion.Controls.OfType<ComboBox>())
            {
                comBoXX.Items.Clear();
                CargarColumnas(comBoXX);
            }
        }

        private void cmbdescripcion_Enter(object sender, EventArgs e)
        {
            if (cmbdescripcion.BackColor == Color.Red)
            {
                cmbdescripcion.BackColor = Color.White;
            }
        }

        private void cmbcodigodebarra_Enter(object sender, EventArgs e)
        {
            if (cmbcodigodebarra.BackColor == Color.Red)
            {
                cmbcodigodebarra.BackColor = Color.White;
            }
        }

        private void cmbidmedida_Enter(object sender, EventArgs e)
        {
            if (cmbidmedida.BackColor == Color.Red)
            {
                cmbidmedida.BackColor = Color.White;
            }
        }

        private void cmbpreciocompra_Enter(object sender, EventArgs e)
        {
            if (cmbpreciocompra.BackColor == Color.Red)
            {
                cmbpreciocompra.BackColor = Color.White;
            }
        }

        private void cmbiva_Enter(object sender, EventArgs e)
        {
            if (cmbiva.BackColor == Color.Red)
            {
                cmbiva.BackColor = Color.White;
            }
        }

        private void cmbprecioventa_Enter(object sender, EventArgs e)
        {
            if (cmbprecioventa.BackColor == Color.Red)
            {
                cmbprecioventa.BackColor = Color.White;
            }
        }

        private void cmbstock_Enter(object sender, EventArgs e)
        {
            if (cmbstock.BackColor == Color.Red)
            {
                cmbstock.BackColor = Color.White;
            }
        }

        private void cmbstock_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbpreciomayorista_Enter(object sender, EventArgs e)
        {
            if (cmbpreciomayorista.BackColor == Color.Red)
            {
                cmbpreciomayorista.BackColor = Color.White;
            }
        }

        private void cmbbajostock_Enter(object sender, EventArgs e)
        {
            if (cmbbajostock.BackColor == Color.Red)
            {
                cmbbajostock.BackColor = Color.White;
            }
        }

        private async void btnimportar_Click(object sender, EventArgs e)
        {
            bool HaycontrolVacio = false;
            foreach (ComboBox control in gimportacion.Controls.OfType<ComboBox>().Where(x => x != cmbiddepartamento && x != cmbidmedida && x.Text == string.Empty))
            {
                HaycontrolVacio = true;
                control.BackColor = Color.Red;
            }
            if (HaycontrolVacio == true)
            {
                MessageBox.Show("Verifique Rellenar los Campos Obligatorios", "Atencion Usuario", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (cmbiddepartamento.Text == string.Empty)
            {
                cmbiddepartamento.Text = "null";

            }
            if (cmbidmedida.Text == string.Empty)
            {
                cmbidmedida.Text = "null";
            }
            lblimportando.Visible = true;
            timerImportacion.Enabled = true;
            using (var ImportarProductos = new Clases.Consultas.CProductos())
            {
               await ImportarProductos.ImportarExcel(lblimportando,lblimportados,lblrechazados, Ruta, cmbhojas.Text.Trim(), cmbdescripcion.Text.Trim()
                    ,cmbcodigodebarra.Text.Trim(), cmbidmedida.Text.Trim(), cmbpreciocompra.Text.Trim(), cmbiva.Text.Trim(), cmbprecioventa.Text.Trim(),
                    cmbstock.Text.Trim(), cmbpreciomayorista.Text.Trim(), cmbbajostock.Text.Trim(), cmbiddepartamento.Text.Trim()
                    );
            }
            timerImportacion.Enabled = false;
            //lblimportados.Text = Importados.ToString();
           // lblrechazados.Text = Fallados.ToString();
        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void timerImportacion_Tick(object sender, EventArgs e)
        {
            if (lblimportando.Visible == true)
                lblimportando.Visible = false;
            else
                lblimportando.Visible = true;
        }
    }
}
