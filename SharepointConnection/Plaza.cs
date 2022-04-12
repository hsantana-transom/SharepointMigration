using System;
using System.Collections.Generic;
using System.Text;

namespace SharepointConnection
{
    class Plaza
    {
        public int IdPlaza;
        public int IdRegion;
        public string cPlaza;
        public string nombrePlaza;

        public Plaza(int id, int idr, string cP, string nombre)
        {
            IdPlaza = id;
            IdRegion = idr;
            cPlaza = cP;
            nombrePlaza = nombre;
        }
    }
}
