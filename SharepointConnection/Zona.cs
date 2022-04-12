using System;
using System.Collections.Generic;
using System.Text;

namespace SharepointConnection
{
    class Zona
    {
        public int IdZona;
        public string nombreZona;
        public string activo;
        public Zona(int id, string nombre, string act)
        {
            IdZona = id;
            nombreZona = nombre;
            activo = act;
        }
    }
}
