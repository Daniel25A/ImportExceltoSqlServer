using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Usings 
using System.Windows.Forms;
using LinqToExcel;
namespace Punto_de_Venta.Clases.Consultas
{
   public enum CamposProductos
    {
        Descripcion,
        IDMedida,
        Stock,
        PrecioCompra,
        PrecioVenta,
        PrecioMayorista,
        IVA,
        Departamento,
        StockMinimo,
        CodigoBarras,
        IncluirIva,
        URlImagen,
    }
   public class CProductos:IDisposable
    {
       private List<string> hojas;
       private List<string> Columnas;
       private List<Object> Productos;

       public static Dictionary<string, decimal> ControldeStock = new Dictionary<string, decimal>();
       public Dictionary<CamposProductos, object> ValoresaGuardar { get; set; }
           
       public List<string> GetHojas(String Ruta = "")
       {
           if (Ruta == string.Empty) return null;

           using (var libro = new ExcelQueryFactory(Ruta))
           {
               hojas = libro.GetWorksheetNames().OrderBy(x => x).ToList<string>();
           }
           if (hojas == null) return null;
           else
               return hojas;
       }
       public CProductos()
       {
           ValoresaGuardar = new Dictionary<CamposProductos, object>();
       }
       public List<string> GetColumnas(String Ruta = "", String Hoja="")
       {
           if (Ruta == string.Empty || Hoja == string.Empty) return null;
           using (var LibroExcel = new ExcelQueryFactory(Ruta))
           {
               Columnas = (from name in LibroExcel.GetColumnNames(Hoja) select name).ToList<string>();
           }
           if (Columnas == null) return null;
           else
               return Columnas;
       }
       public async Task ImportarExcel (Label lblimportando,Label lblimportados, Label lblrechazados,string Ruta, string Libro
           ,string CMNAME,
           string CMCODIGO,
           string CMIDMEDIDA,
           string CMPCOMPRA,
           string CMIVA,
           string CMPVENTA,
           string CMSTOCK,
           string CMPMAYORISTA,
           string CMBAJOSTOCK,
           string CMIDDEPARTAMENTO
           )
       {
           int Fallidos = 0;
           int Importados = 0;
          
             await  Task.Run(() => {
                 try
                 {
                     using (var storeData = new DataBaseControllerDataContext())
                     {
                         using (var LibroExcel = new ExcelQueryFactory(Ruta))
                         {

                             var Valores = from valor in LibroExcel.Worksheet(Libro)
                                           let values = new
                                           {
                                               idMedida = CMIDMEDIDA == "null" ? 1 : valor[CMIDMEDIDA].Cast<int>(),
                                               idDepartamento = CMIDDEPARTAMENTO == "null" ? 1 : valor[CMIDDEPARTAMENTO].Cast<int>(),
                                               pNombre = valor[CMNAME].Cast<string>(),
                                               pCodigo = valor[CMCODIGO].Cast<string>(),
                                               pIva = valor[CMIVA].Cast<float>(),
                                               pCompra = valor[CMPCOMPRA].Cast<decimal>(),
                                               pVenta = valor[CMPVENTA].Cast<decimal>(),
                                               pMayorista = valor[CMPMAYORISTA].Cast<decimal>(),
                                               pStock = valor[CMSTOCK].Cast<int>(),
                                               pMinimo = valor[CMBAJOSTOCK].Cast<int>()
                                           }
                                           select values;
                             foreach (var x in Valores)
                             {
                                 var temp = new TABProductos();
                                 if (storeData.GetTable<TABProductos>().Any(c => c.codigoBarras == x.pCodigo)) { Fallidos++; continue; }
                                 temp.nombreProducto = x.pNombre;
                                 temp.codigoBarras = x.pCodigo;
                                 temp.idUnidadMedida = x.idMedida;
                                 temp.idDepartamento = x.idDepartamento;
                                 temp.precioUnitarioCompra = x.pCompra;
                                 temp.precioUnitarioMinuista = x.pVenta;
                                 temp.precioUnitarioMayorista = x.pMayorista;
                                 temp.IVA = x.pIva;
                                 temp.cantidadStock = x.pStock;
                                 temp.cantidadminimaStock = x.pMinimo;
                                 storeData.TABProductos.InsertOnSubmit(temp);
                                 storeData.SubmitChanges();
                                 temp = null;
                                 Importados++;
                             }
                             Valores = null;
                         }
                     }
                 }
                 catch (FormatException)
                 {
                     MessageBox.Show("Verifique que esta Agregando los campos Correctamente", "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 }
                 catch (Exception)
                 {
                     MessageBox.Show("Ocurrio un Error al importar los datos de Excel", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 }
               });
           lblimportados.Text = Importados.ToString();
           lblrechazados.Text = Fallidos.ToString();
           lblimportando.Visible = false;
       }

       public void RegistrarProducto(IDictionary<CamposProductos, object> Campos)
       {
           if (Campos.Count < 12 || Campos.Count == 0)
           {
               MessageBox.Show("Asegurese de Rellenar todos los Parametros Requeridos");
               return;
           }
           try
           {
               var Valores = Campos as Dictionary<CamposProductos, object>;
               if (Valores == null) throw new InvalidOperationException();
               using (var storeData = new DataBaseControllerDataContext())
               {
                   //--- SI LA CLAVE NO EXISTE DEBE LANZAR UNA EXCEPCION DE TIPO  KeyNotFoundException LA CUAL YA ESTA CONTROLADA
                   var TempProductos = new TABProductos();
                   var TempIva = new TabIncluirIva();
                   TempProductos.nombreProducto = (String)Valores[CamposProductos.Descripcion];
                   TempProductos.idDepartamento = Convert.ToInt32(Valores[CamposProductos.Departamento]);
                   TempProductos.idUnidadMedida = Convert.ToInt32(Valores[CamposProductos.IDMedida]);
                   TempProductos.IVA = 0D;
                   TempProductos.precioUnitarioCompra = decimal.Parse(Valores[CamposProductos.PrecioCompra].ToString());
                   TempProductos.precioUnitarioMinuista = decimal.Parse(Valores[CamposProductos.PrecioVenta].ToString());
                   TempProductos.precioUnitarioMayorista = decimal.Parse(Valores[CamposProductos.PrecioMayorista].ToString());
                   TempProductos.cantidadStock = Decimal.Parse(Valores[CamposProductos.Stock].ToString());
                   TempProductos.cantidadminimaStock = Decimal.Parse(Valores[CamposProductos.StockMinimo].ToString());
                   TempProductos.codigoBarras = (String)Valores[CamposProductos.CodigoBarras];
                   TempProductos.PathImagen = (String)Valores[CamposProductos.URlImagen];
                   TempIva.CodigodeBarras = (String)Valores[CamposProductos.CodigoBarras];
                   TempIva.IncluirIva = Convert.ToInt32((Boolean)Valores[CamposProductos.IncluirIva]);
                   TempIva.IVA = double.Parse(Valores[CamposProductos.IVA].ToString());
                   storeData.TABProductos.InsertOnSubmit(TempProductos);
                   storeData.TabIncluirIva.InsertOnSubmit(TempIva);
                   storeData.SubmitChanges();
                   TempProductos = null;
                   TempIva = null;
               }
               Valores = null;
           }
           catch (KeyNotFoundException)
           {
              MessageBox.Show("El Valor al que Intentas acceder no Existe", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
           }
           catch (InvalidCastException)
           {
               MessageBox.Show("Aseguere de que no esta ingresando valores invalidos", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
           }
           catch (InvalidOperationException)
           {
               MessageBox.Show("Usted ha Realizado una Operación Invalida\nContacte a su Proveedor si el error sigue apareciendo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
           }
       }
       public void EliminarProducto(String CodigoBarras)
       {
           try
           {
               using (var storeData = new DataBaseControllerDataContext())
               {
                   var TempProducto = storeData.GetTable<TABProductos>().Where(x => x.codigoBarras == CodigoBarras.Trim()).FirstOrDefault();
                   storeData.TABProductos.DeleteOnSubmit(TempProducto);
                   var TempIva = storeData.GetTable<TabIncluirIva>().FirstOrDefault(x => x.CodigodeBarras == CodigoBarras);
                   storeData.TabIncluirIva.DeleteOnSubmit(TempIva);
                   storeData.SubmitChanges();
                   MessageBox.Show(string.Format("El Producto {0} Ha Sido Eliminado", CodigoBarras), "Operacion Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   TempProducto = null;
               }
           }
           catch (Exception)
           {

               MessageBox.Show("Error en Tiempo de Ejecucion");
           }
       }
       public void ActualizarProducto(TABProductos Entity,TabIncluirIva Entity2)
       {
           try
           {
               using (var storeData = new DataBaseControllerDataContext())
               {
                   var UpdateProd = (from xPro in storeData.GetTable<TABProductos>() where xPro.codigoBarras == Entity.codigoBarras select xPro).FirstOrDefault();
                   UpdateProd.nombreProducto = Entity.nombreProducto;
                   UpdateProd.idUnidadMedida = Entity.idUnidadMedida;
                   UpdateProd.idDepartamento = Entity.idDepartamento;
                   UpdateProd.cantidadStock = Entity.cantidadStock;
                   UpdateProd.cantidadminimaStock = Entity.cantidadminimaStock;
                   UpdateProd.precioUnitarioCompra = Entity.precioUnitarioCompra;
                   UpdateProd.precioUnitarioMinuista = Entity.precioUnitarioMinuista;
                   UpdateProd.precioUnitarioMayorista = Entity.precioUnitarioMayorista;
                   UpdateProd.PathImagen = Entity.PathImagen;
                   var UpdateIVA = storeData.GetTable<TabIncluirIva>().Where(x => x.CodigodeBarras == Entity.codigoBarras).FirstOrDefault();
                   UpdateIVA.IncluirIva = Entity2.IncluirIva;
              //     UpdateIVA.IVA = Entity2.IVA;
                   storeData.SubmitChanges();
               }
           }
           catch (Exception)
           {
               
              MessageBox.Show("Error en Tiempo de Ejecucion");
           }
       }
       public void EliminarProducto(TABProductos Entity)
       {
           try
           {
               using (var storeData = new DataBaseControllerDataContext())
               {
                   var TempTableProductos = storeData.GetTable<TABProductos>().FirstOrDefault(c => c.codigoBarras == Entity.codigoBarras);
                   storeData.TABProductos.DeleteOnSubmit(TempTableProductos);
                   storeData.SubmitChanges();
               }
           }
           catch (Exception ex)
           {
               MessageBox.Show("Error en " + ex.TargetSite);
           }
       }
       public List<Object> CargarProductos()
       {
           try
           {
               using (var storeData = new DataBaseControllerDataContext())
               {
                   Productos = (from producto in storeData.GetTable<TABProductos>()
                                join departamento in storeData.GetTable<TabDepartamento>()
                                on producto.idDepartamento equals departamento.id
                                join medida in storeData.GetTable<TabUnidadMedida>()
                                on producto.idUnidadMedida equals medida.id
                                select new { DESCRIPCION=producto.nombreProducto,
                                COSTO=producto.precioUnitarioCompra,
                                VENTA=producto.precioUnitarioMinuista,
                                MAYORISTA=producto.precioUnitarioMayorista,
                                CODIGODEBARRA=producto.codigoBarras,
                                DEPARTAMENTO=departamento.Descripcion,
                                UMEDIDA=medida.Descripcion
                                }
                                    ).ToList<Object>();
                   if (Productos == null) throw new Exception();
               }
           }
           catch (Exception)
           {
               MessageBox.Show("Error");
               Productos = null;
           }
           return Productos;
       }
       public static void ControlarStock(String Codigo="", decimal Cantidad=0,Boolean Vaciar=false)
       {
           try
           {
               if (Vaciar)
               {
                   ControldeStock.Clear();
                   return;
               }
               if (ControldeStock.ContainsKey(Codigo))
                   ControldeStock[Codigo] = ControldeStock[Codigo] + Cantidad;
               else
                   ControldeStock.Add(Codigo, Cantidad);
           }
           catch (Exception)
           {
               MessageBox.Show("Ocurrio un Error al Controlar el Stock", "Atencion Usuario", MessageBoxButtons.OK, MessageBoxIcon.Error);
           }
       }
       public static Boolean VerificarStockBajo(String CodigodeBarra, ref TABProductos Producto)
       {
           Boolean TenemosStockBajo = false;
           try
           {
               if ((Producto.cantidadStock - ControldeStock[CodigodeBarra]) <= Producto.cantidadminimaStock)
               {
                   TenemosStockBajo = true;
               }
           }
           catch (Exception)
           {

               MessageBox.Show("Error al controlar el stock bajo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
               TenemosStockBajo = false;
           }
           return TenemosStockBajo;
       }
       public void Dispose()
       {
           if (hojas != null)
               hojas = null;
           if (Columnas != null)
               Columnas = null;
           if (ValoresaGuardar != null)
               ValoresaGuardar = null;
           if (Productos != null)
               Productos = null;
       }
    }
}
