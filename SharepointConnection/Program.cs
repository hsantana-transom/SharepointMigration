using System;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.IO;


namespace SharepointConnection  
{
    class Program
    {
        SqlConnection conn = new SqlConnection();
        List<Zona> ListaZonas= new List<Zona>();
        List<Region> ListaRegiones= new List<Region>();
        List<Plaza> ListaPlazas= new List<Plaza>();
        List<Tienda> ListaTiendas= new List<Tienda>();

        List<Zona> ListaShareZonas = new List<Zona>();
        List<Region> ListaShareRegiones = new List<Region>();
        List<Plaza> ListaSharePlazas = new List<Plaza>();
        List<Tienda> ListaShareTiendas = new List<Tienda>();
        
        static void Main(string[] args)
        {

            string siteUrl = "https://wrmdp.sharepoint.com/sites/DemoSite/";
            Program s = new Program();


            string selectedOption = s.menu();

            if (selectedOption != "1" && selectedOption != "2")
                Console.WriteLine("Opción no válida");
            if(selectedOption.Trim()=="1")
                s.EncryptFile();
            if (selectedOption.Trim() == "2")
            {
                string pass=s.DecryptPass();

                if (pass != "Password Incorrecto")
                {

                    s.SQLConnection("Zonas");
                    s.SQLConnection("Regiones");
                    s.SQLConnection("PlazasEmpoderadas");
                    s.SQLConnection("Tiendas");
                    using (var cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, "59f59021-604f-40f1-b88f-344bc3c63469", "wrmdp.onmicrosoft.com", @"C:\Windows\System32\hexaware.pfx", pass))
                    {
                        cc.Load(cc.Web, p => p.Title);
                        cc.ExecuteQuery();
                        Console.WriteLine("Succesful Connection: " + cc.Web.Title);
                        s.SharepointData(cc, "Zonas");
                        s.SharepointData(cc, "Regiones");
                        s.SharepointData(cc, "Plazas");
                        s.SharepointData(cc, "Tiendas");

                        s.checkDuplicates("Zonas");
                        s.checkDuplicates("Regiones");
                        s.checkDuplicates("Plazas");
                        s.checkDuplicates("Tiendas");


                        s.SharepointSave(cc, "Zonas");
                        s.SharepointSave(cc, "Regiones");
                        s.SharepointSave(cc, "Plazas");
                        s.SharepointSave(cc, "Tiendas");
                    }
                }
            }
            Console.WriteLine("Presione una tecla para terminar ...");
            Console.ReadKey();


        }
        public string menu()
        {
            Console.WriteLine("Selecciona una opcion.");
            Console.WriteLine("1. Guardar Password Encriptado.");
            Console.WriteLine("2. Leer Archivo con password y conectarse a Sharepoint.");
            Console.Write("Opción: ");
            string selectedOption = Console.ReadLine();
            return selectedOption;
        }
        public void EncryptFile()
        {

            Console.WriteLine("Please enter a string to encrypt:");
            string plaintext = Console.ReadLine();
            Console.WriteLine("Please enter a password to use:");
            string password = Console.ReadLine();
            Console.WriteLine("");

            Console.WriteLine("Your encrypted string is:");
            string encryptedstring = StringCipher.Encrypt(plaintext, password);
            Console.WriteLine(encryptedstring);
            Console.WriteLine("");

           
        }
        public string DecryptPass()
        {
            Console.WriteLine("Please enter a password to decrypt:");
            string password = Console.ReadLine();

            string encryptedstring=System.IO.File.ReadAllText("C:\\Users\\usuario\\Documents\\hexaware\\passAzure.txt");
            Console.WriteLine("Your decrypted string is:");
            string decryptedstring = StringCipher.Decrypt(encryptedstring, password);
            return decryptedstring; 
           

        }
        /**
         *  Método que se conecta al sitio de Sharpeoint para guardar la informacion de las listas de SQL
         *  
         *  Recibe como parametros el ClientContext para hacer la conexion con el sitio y el nombre de la lista de Sharepoint
         *  donde se almacenaran los datos
         */
        private void SharepointSave(ClientContext cc,string shList)
        {
            List listSh = cc.Web.Lists.GetByTitle(shList);
            ListItemCreationInformation itemCreationInfo = new ListItemCreationInformation();
            ListItem itemToInsert;// = listSh.AddItem(itemCreationInfo);

            int retryAttempts = 0;
            int Attempts = 5;
            switch (shList)
            {
                case "Zonas":
                    ListaZonas.ForEach(z =>
                    {
                        itemToInsert = listSh.AddItem(itemCreationInfo);
                        retryAttempts = 0;
                        itemToInsert["IdZona"] = z.IdZona;
                        itemToInsert["Nombre"] = z.nombreZona;
                        itemToInsert["Activo"] = z.activo == "1" ? true : false;
                        itemToInsert.Update();

                        try
                        {
                            Console.WriteLine("Insertando Zona: " + z.nombreZona);
                            cc.ExecuteQuery();
                            
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                      
                            while (retryAttempts <= Attempts)
                            {
                                retryAttempts++;
                                Console.WriteLine("ReIntento " + retryAttempts + "...");
                                Thread.Sleep(retryAttempts * 10000);
                                try
                                {
                                    cc.ExecuteQuery();
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e);
                                }
                            }

                        }

                    });                    
                break;
                case "Regiones":
                    ListaRegiones.ForEach(r =>
                    {
                        itemToInsert = listSh.AddItem(itemCreationInfo);
                        retryAttempts = 0;
                        itemToInsert["IdRegion"] = r.Idregion;
                        itemToInsert["IdZona"] = r.IdZona;
                        itemToInsert["Nombre"] = r.nombreRegion;
                        itemToInsert.Update();

                        try
                        {
                            Console.WriteLine("Insertando Region: " + r.nombreRegion);
                            cc.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);

                            while (retryAttempts <= Attempts)
                            {
                                retryAttempts++;
                                Console.WriteLine("ReIntento " + retryAttempts + "...");
                                Thread.Sleep(retryAttempts * 10000);
                                itemToInsert.Update();
                                try
                                {
                                    cc.ExecuteQuery();
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e);
                                }
                            }

                        }

                    });
                break;
                case "Plazas":
                    ListaPlazas.ForEach(p =>
                    {
                        itemToInsert = listSh.AddItem(itemCreationInfo);
                        retryAttempts = 0;
                        itemToInsert["IdPlaza"] = p.IdPlaza;
                        itemToInsert["IdRegion"] = p.IdRegion;
                        itemToInsert["cPlaza"] = p.cPlaza;
                        itemToInsert["Nombre"] = p.nombrePlaza;
                        itemToInsert.Update();
                        try
                        {
                            Console.WriteLine("Insertando Plaza: " + p.nombrePlaza);
                            
                            cc.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);

                            while (retryAttempts <= Attempts)
                            {
                                retryAttempts++;
                                Console.WriteLine("ReIntento " + retryAttempts + "...");
                                Thread.Sleep(retryAttempts * 10000);
                                itemToInsert.Update();
                                try
                                {
                                    cc.ExecuteQuery();
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e);
                                }
                            }

                        }

                    });
                break;
                case "Tiendas":
                    ListaTiendas.ForEach(t =>
                    {
                        itemToInsert = listSh.AddItem(itemCreationInfo);
                        retryAttempts = 0;
                        itemToInsert["IdTienda"] = t.IdTienda;
                        itemToInsert["IdPlaza"] = t.IdPlaza;
                        itemToInsert["cTienda"] = t.cTienda;
                        itemToInsert["Nombre"] = t.nombreTienda;
                        itemToInsert.Update();

                        try
                        {
                            Console.WriteLine("Insertando Tienda: " + t.nombreTienda);
                            cc.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);

                            while (retryAttempts <= Attempts)
                            {
                                retryAttempts++;
                                Console.WriteLine("ReIntento " + retryAttempts + "...");
                                Thread.Sleep(retryAttempts * 10000);
                                itemToInsert.Update();
                                try
                                {
                                    cc.ExecuteQuery();
                                    break;
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e);
                                }
                            }

                        }

                    });
                break;
            }

        }

        /**
         * Método que se conecta a las listas de Sharepoint para obterner los datos
         * 
         * Recibe como parmetros el clientContext para la conexion al sitio y el nombre de la lista donde 
         * se obtendran los datos
         */
        private void SharepointData(ClientContext cc, string listName)
        {
            List targetList = cc.Web.Lists.GetByTitle(listName);
            CamlQuery oquery = CamlQuery.CreateAllItemsQuery();
            ListItemCollection ocollection = targetList.GetItems(oquery);
            cc.Load(ocollection);
            try
            {
                cc.ExecuteQuery();

                foreach (ListItem item in ocollection)
                {
                    switch (listName)
                    {
                        case "Zonas":
                            Zona z = new Zona(Convert.ToInt32(item["IdZona"]),item["Nombre"].ToString(),item["Activo"].ToString()=="True"? "1" : "0");
                            ListaShareZonas.Add(z);

                        break;
                        case "Regiones":
                            Region r = new Region(Convert.ToInt32(item["IdRegion"]), Convert.ToInt32(item["IdZona"]),item["Nombre"].ToString());
                            ListaShareRegiones.Add(r);
                        break;
                        case "Plazas":
                            Plaza p = new Plaza(Convert.ToInt32(item["IdPlaza"]),Convert.ToInt32(item["IdRegion"]),item["cPlaza"].ToString(),item["Nombre"].ToString());
                            ListaSharePlazas.Add(p);
                        break;
                        case "Tiendas":
                            Tienda t = new Tienda(Convert.ToInt32(item["IdTienda"]), Convert.ToInt32(item["IdPlaza"]), item["cTienda"].ToString(), item["Nombre"].ToString());
                            ListaShareTiendas.Add(t);
                        break;
                    }
                   // Console.WriteLine(item["Title"].ToString());

                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        /**
            Método que conecta a la base de datos y llena las listas
            
            Recibe como parametro el nombre de la lista a consultar(Zonas,Regiones,Plazas,Tiendas)
        */
        private void SQLConnection(string sqlTable)
        {
            conn.ConnectionString =
              "Data Source=LENOVO-PC;" +
              "Initial Catalog=OXXO;" +
              "Integrated Security=SSPI";
            conn.Open();
            SqlCommand command;
            string sql = null;
            try
            {
                switch (sqlTable)
                {
                    case "Zonas":
                        sql = "SELECT id_zona,zona,activo FROM " + sqlTable;
                    break;
                    case "Regiones":
                        sql = "SELECT id_region,idzona,regionNombre FROM " + sqlTable;
                    break;
                    case "PlazasEmpoderadas":
                        sql = "SELECT idplaza,idregion,crplaza,plazaNombre FROM " + sqlTable;
                        break;
                    case "Tiendas":
                        sql = "SELECT idtienda,idplaza,crtienda,nombre_tienda FROM " + sqlTable;
                    break;
                }
                        
                command = new SqlCommand(sql, conn);
                SqlDataReader reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    switch (sqlTable)
                    {
                        case "Zonas":
                            Zona z = new Zona((int)reader[0],(string)reader[1],(string)reader[2]);
                            ListaZonas.Add(z);
                            
                        break;
                        case "Regiones":
                            Region r = new Region((int)reader[0],(int)reader[1],(string)reader[2]);
                            ListaRegiones.Add(r);
                        break;
                        case "PlazasEmpoderadas":
                            Plaza p = new Plaza((int)reader[0], (int)reader[1], (string)reader[2], (string)reader[3]);
                            ListaPlazas.Add(p);
                        break;
                        case "Tiendas":
                            Tienda t = new Tienda((int)reader[0], (int)reader[1], (string)reader[2], (string)reader[3]);
                            ListaTiendas.Add(t);
                        break;
                    }
                }

                conn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        
        /*
            Método para verificar items existentes o cambios en los items.
            - Si el Item ya existe en las listas de Sharepoint se elimina de la Lista de los items de SQL
            - Si el item existe y cuenta con algun cambio se guardara en archivo CSV

            Recibe como parámetro el nombre de la lista a verificar(Zonas,Regiones,Plazas,Tiendas)

        */
        private void checkDuplicates(string listName)
        {
            DataTable table = new DataTable();
            CSVUtlity file = new CSVUtlity();
            switch (listName)
            {
                case "Zonas":
                    table.Columns.Add("ID", typeof(int));
                    table.Columns.Add("Sharepoint Nombre Zona", typeof(string));
                    table.Columns.Add("Sharepoint Activo", typeof(string));
                    table.Columns.Add("SQL Nombre Zona", typeof(string));
                    table.Columns.Add("SQL Activo", typeof(string));

                    if (ListaZonas.Count > 0 && ListaShareZonas.Count > 0)
                    {
                        List<Zona> listaAux = new List<Zona>(ListaZonas);
                        
                        listaAux.ForEach(z=> {
                            var item = ListaShareZonas.Where(sItem => sItem.IdZona == z.IdZona);
                            foreach (var zona in item)
                            {
                                if (zona.nombreZona != z.nombreZona || zona.activo != z.activo)
                                {
                                    table.Rows.Add(z.IdZona, zona.nombreZona, zona.activo, z.nombreZona, z.activo);
                                }
                                ListaZonas.RemoveAt(ListaZonas.FindIndex(zone => zone.IdZona == z.IdZona));

                            }
                           
                        });
                    }
                    
                    file.ToCSV(table, "C:\\CSVFiles\\ZonasCambios");

                break;
                case "Regiones":
                    table.Columns.Add("ID", typeof(int));
                    table.Columns.Add("Sharepoint Nombre Región", typeof(string));
                    table.Columns.Add("Sharepoint IdZona", typeof(string));
                    table.Columns.Add("SQL Nombre Región", typeof(string));
                    table.Columns.Add("SQL IdZona", typeof(string));
                    if (ListaRegiones.Count > 0 && ListaShareRegiones.Count > 0)
                    {
                        List<Region> listaAux = new List<Region>(ListaRegiones);
                        listaAux.ForEach(r =>
                        {
                            var item = ListaShareRegiones.Where(sItem => sItem.Idregion == r.Idregion);
                            foreach (var region in item)
                            {
                                if (region.IdZona != r.IdZona || region.nombreRegion != r.nombreRegion)
                                {
                                    table.Rows.Add(r.Idregion, region.nombreRegion, region.IdZona, r.nombreRegion, r.IdZona);
                                }
                                ListaRegiones.RemoveAt(ListaRegiones.FindIndex(reg=> reg.Idregion == r.Idregion));

                            }
                        });
                    }
                    file.ToCSV(table, "C:\\CSVFiles\\RegionesCambios");
                break;

                case "Plazas":
                    table.Columns.Add("ID", typeof(int));
                    table.Columns.Add("Sharepoint Nombre Plaza", typeof(string));
                    table.Columns.Add("Sharepoint cv Plaza", typeof(string));
                    table.Columns.Add("Sharepoint Id Región", typeof(string));
                    table.Columns.Add("SQL Nombre Plaza", typeof(string));
                    table.Columns.Add("SQL cv Plaza", typeof(string));
                    table.Columns.Add("SQL Id Región", typeof(string));
                    if (ListaPlazas.Count > 0 && ListaSharePlazas.Count > 0)
                    {
                        List<Plaza> listaAux = new List<Plaza>(ListaPlazas);
                        listaAux.ForEach(p =>
                        {
                            var item = ListaSharePlazas.Where(sItem => sItem.IdPlaza == p.IdPlaza);
                            foreach (var plaza in item)
                            {
                                if (plaza.IdRegion != p.IdRegion || plaza.nombrePlaza != p.nombrePlaza || plaza.cPlaza != p.cPlaza)
                                {
                                    table.Rows.Add(p.IdPlaza, plaza.nombrePlaza, plaza.cPlaza, plaza.IdRegion, p.nombrePlaza, p.cPlaza, p.IdRegion);
                                }
                                ListaPlazas.RemoveAt(ListaPlazas.FindIndex(pla => pla.IdPlaza == p.IdPlaza));
                            }
                        });
                    }
                    file.ToCSV(table, "C:\\CSVFiles\\PlazasCambios");
                break;
                case "Tiendas":
                    table.Columns.Add("ID", typeof(int));
                    table.Columns.Add("Sharepoint Nombre Tienda", typeof(string));
                    table.Columns.Add("Sharepoint cv Tienda", typeof(string));
                    table.Columns.Add("Sharepoint Id Plaza", typeof(string));
                    table.Columns.Add("SQL Nombre Tienda", typeof(string));
                    table.Columns.Add("SQL cv Tienda", typeof(string));
                    table.Columns.Add("SQL Id Plaza", typeof(string));
                    if (ListaTiendas.Count > 0 && ListaShareTiendas.Count > 0)
                    {
                        List<Tienda> listaAux = new List<Tienda>(ListaTiendas);

                        listaAux.ForEach(t =>
                        {
                            var item = ListaShareTiendas.Where(sItem => sItem.IdTienda == t.IdTienda);
                            foreach (var tienda in item)
                            {
                                if (tienda.IdPlaza != t.IdPlaza || tienda.nombreTienda != t.nombreTienda || tienda.cTienda != t.cTienda)
                                {
                                    table.Rows.Add(t.IdTienda, tienda.nombreTienda, tienda.cTienda, tienda.IdPlaza, t.nombreTienda, t.cTienda, t.IdPlaza);
                                }
                                ListaTiendas.RemoveAt(ListaTiendas.FindIndex(ti => ti.IdTienda == t.IdTienda));
                            }
                        });
                    }
                    file.ToCSV(table, "C:\\CSVFiles\\TiendasCambios");
                break;
            }

        }
    }
}
