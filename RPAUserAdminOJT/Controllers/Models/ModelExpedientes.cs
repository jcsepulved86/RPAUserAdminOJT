using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAUserAdminOJT.Controllers.Models
{
    public class ModelExpedientes
    {
        public string documento { get; set; }
        public string nombre_completo { get; set; }
        public string ceco { get; set; }
        public string nombre_ceco { get; set; }
        public string cod_pcrc { get; set; }
        public string nombre_pcrc { get; set; }
        public string id_dp_estados { get; set; }
        public string tipo_estado { get; set; }
        public string id_dp_cargos { get; set; }
        public string nombre_cargo { get; set; }
        public string documento_jefe { get; set; }
        public string nombre_jefe { get; set; }
        public string email_jefe { get; set; }
        public string fecha_modificacion { get; set; }
        public string id_dp_solicitudes { get; set; }


        public static ArrayList candidatForma = new ArrayList();
        public static ArrayList candidatOJT = new ArrayList();
        public static ArrayList candidatRechz = new ArrayList();
        public static ArrayList candidatBeginner = new ArrayList();
        public static ArrayList candidatMovement = new ArrayList();

    }
}
