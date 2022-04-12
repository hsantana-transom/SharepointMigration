using System;
using System.Collections.Generic;
using System.Text;

namespace SharepointConnection
{
    class Tienda
    {
        public int IdTienda;
        public int IdPlaza;
        public string cTienda;
        public string nombreTienda;

        public Tienda(int id, int idP, string cT, string nombre)
        {
            IdTienda = id;
            IdPlaza = idP;
            cTienda = cT;
            nombreTienda = nombre;
        }
    }
}
