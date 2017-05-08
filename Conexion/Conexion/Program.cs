using System;
using System.IO;
using SAPbobsCOM;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace Conexion
{
    class Program
    {
        #region Global Variables 
        //variavbles del programa 
        static SAPbobsCOM.Company oCompany = null;
        static string sRuta;
        #endregion
        static void Main(string[] args)
        {
            //Se define la ruta de ejecución del .EXE
            string sRutaCompleta = Process.GetCurrentProcess().MainModule.FileName;
            sRuta = Path.GetDirectoryName(sRutaCompleta);
            //Conexión           
            Conexion();

            //Metodos implementados
            oCompany.StartTransaction();
            addBP();
            updateBP();
            addOrder();
            oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
            //addInvoiceFromOrder();
            //Al terminar el programa se debe desconectar la conexion
            oCompany.Disconnect();
            GC.Collect();
        }
        static void Conexion()
        {
            //Inicia conexión DI API
            try
            { 
                //datos básicos de conexion
                oCompany = new SAPbobsCOM.Company();
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                oCompany.Server = "gdljvaldovi";
                oCompany.UserName = "manager";
                oCompany.Password = "manager";
                oCompany.CompanyDB = "SBODemoMX";

                int iRetCode = oCompany.Connect();
                if (iRetCode == 0)
                {
                    Log("EXITO", "Conexión exitosa al DI API, conectado: " + oCompany.UserName);
                } else
                {
                    Log("ERROR", "No se pudo conectar al DI API - " + oCompany.GetLastErrorDescription());
                }

            }
            catch (Exception e)
            {
                Log("ERROR", "No se pudo conectar al DI API - " + oCompany.GetLastErrorDescription() + " - " + e.ToString());
            }

            //Finaliza conexión DI API
        }
        static void Log(string sEstatus, string sMsg, [CallerMemberName] string sMethod = "", [CallerLineNumber] int iLine = 0)
        {
            try
            {
                //Crea la carpeta de Logs y Estatus
                Directory.CreateDirectory(sRuta + "\\Logs\\" + sEstatus);
                //Crea el archivo de Logs y concatena los mensajes
                using (StreamWriter w = File.AppendText(sRuta + "\\Logs\\" + sEstatus + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt"))
                {
                    w.WriteLine("Log Entry: {0} {1} - {2}({3})", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString(), sMethod, iLine);
                    w.WriteLine("   {0}", sMsg);
                    w.WriteLine("---------------------------------------------------------------------------------------------------------------");
                }
                if (sEstatus != "WARNING") //Si el estatus es Warning, no lo imprime en consola
                {
                    Console.WriteLine(sEstatus + " - " + sMsg);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR - " + e.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        static void addBP()
        {
            //objeto para socios de negocio
            SAPbobsCOM.BusinessPartners oBP;
            try
            {
                //Información del BP
                oBP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                oBP.CardCode = "C08154712";
                oBP.CardName = "James Tiberius Kirk II";
                oBP.CardType = BoCardTypes.cCustomer;
                oBP.FederalTaxID = "XAXX010101000";

                int iRetCode = oBP.Add();
                if (iRetCode == 0)
                {
                    Log("EXITO", "se agrego el socio de negocio" + oBP.CardCode);
                }
                else
                {
                    Log("ERROR", oCompany.GetLastErrorDescription());
                }
            }
            catch
            {
                Log("ERROR", oCompany.GetLastErrorDescription());
            }
        }
        static void updateBP()
        {
            SAPbobsCOM.BusinessPartners oBP;
            try
            {
                //primero se valida  que exista en el sistema
                oBP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                if (oBP.GetByKey("C08154711"))
                {
                    Log("EXITO", "Si existe " + oBP.CardName);
                    oBP.City = "GDL";
                    int iRetCode = oBP.Update();
                    if (iRetCode == 0)
                    {
                        Log("EXITO", "se actualizo el usuario " + oBP.CardName);
                    } else
                    {
                        Log("ERROR", oCompany.GetLastErrorDescription());
                    }
                } else
                {
                    Log("ERROR", "NO existe C08154711");

                }

            }
            catch
            {
                Log("ERROR", oCompany.GetLastErrorDescription());
            }
        }
        static void addOrder()
        {
            SAPbobsCOM.Documents oOrder;
            try
            {
                //se llena la información del documento
                oOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                oOrder.CardCode = "C20000";
                oOrder.DocDueDate = DateTime.Now;
                oOrder.Lines.ItemCode = "A00001";
                oOrder.Lines.Quantity = 1;
                int iRetCode = oOrder.Add();
                if (iRetCode == 0)
                {
                    //se obtiene el numero del ultimo documento creado
                    Log("EXITO", "se agrego correctamente " + oCompany.GetNewObjectKey());
                }
                else
                {
                    Log("ERROR", oCompany.GetLastErrorDescription());
                }
            }
            catch
            {
                Log("ERROR", oCompany.GetLastErrorDescription());
            }
        } 
        static void addInvoiceFromOrder()
        {
            // se crea objeto  para la factura
            SAPbobsCOM.Documents oInvoice;
            try
            {
                oInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                oInvoice.DocDate = DateTime.Now;
                oInvoice.CardCode = "C20000";
                //oInvoice.Lines.BaseType = BoObjectTypes.oOrders;

                int iOrderDocNum;
                int.TryParse("123", out iOrderDocNum);
                oInvoice.Lines.ItemCode = "A00001";
                int iRetCode = oInvoice.Add();                
                oInvoice.Lines.Quantity = 1;
                if (iRetCode == 0)
                {
                    //se obtiene el numero del ultimo documento creado
                    Log("EXITO", "se agrego correctamente " + oCompany.GetNewObjectKey());
                }
                else
                {
                    Log("ERROR", oCompany.GetLastErrorDescription());
                }
            }
            catch
            {
                Log("ERROR", oCompany.GetLastErrorDescription());
            }
        }
    }
}
