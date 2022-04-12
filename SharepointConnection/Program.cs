using System;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

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
            s.SQLConnection("Zonas");
            s.SQLConnection("Regiones");
            s.SQLConnection("PlazasEmpoderadas");
            s.SQLConnection("Tiendas");
            using (var cc = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteUrl, "59f59021-604f-40f1-b88f-344bc3c63469", "wrmdp.onmicrosoft.com", @"C:\Windows\System32\hexaware.pfx", "dev2022"))
            {
                cc.Load(cc.Web, p => p.Title);
                cc.ExecuteQuery();
                Console.WriteLine("Succesful Connection: " + cc.Web.Title);
                s.SharepointData(cc, "Zonas");
                s.SharepointData(cc, "Regiones");
                s.SharepointData(cc, "Plazas");
                s.SharepointData(cc, "Tiendas");
                s.SharepointSave(cc, "Zonas");

            }
          
        }
        private void SharepointSave(ClientContext cc,string shList)
        {
            List listSh = cc.Web.Lists.GetByTitle(shList);
            ListItemCreationInformation itemCreationInfo = new ListItemCreationInformation();
            ListItem itemToInsert = listSh.AddItem(itemCreationInfo);

            int retryAttempts = 0;
            int Attempts = 5;
            switch (shList)
            {
                case "Zonas":
                    ListaZonas.ForEach(z =>
                    {
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
                            Zona z = new Zona(Convert.ToInt32(item["IdZona"]),item["Nombre"].ToString(),item["Activo"].ToString());
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
    }
}
