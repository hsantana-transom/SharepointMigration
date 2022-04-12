using System;
using System.Collections.Generic;
using System.Text;

namespace SharepointConnection
{
    class Region
    {
        public int Idregion;
        public int IdZona;
        public string nombreRegion;
        public Region(int id, int idz, string nombre)
        {
            Idregion = id;
            IdZona = idz;
            nombreRegion = nombre;

        }
    }
}
