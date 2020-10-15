using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAUserAdminOJT.Controllers.Models
{
    public class ModelExpedientes
    {
        public string id { get; set; }
        public string nmAsistencia { get; set; }
        public string feSolicitud { get; set; }
        public string dniAsegurado { get; set; }
        public string dsNombreAsegurado { get; set; }
        public string dniSolicitante { get; set; }
        public string dsNombreSolicitante { get; set; }
        public string nmTelefono1 { get; set; }
        public string nmTelefono2 { get; set; }
        public string feCita { get; set; }
        public object dsMunicipioOrigen { get; set; }
        public string dsDepartamentoOrigen { get; set; }
        public string dsDireccionOrigen { get; set; }
        public string dsDetalleDireccionOrigen { get; set; }
        public string dsMunicipioDestino { get; set; }
        public string dsDepartamentoDestino { get; set; }
        public string dsDireccionDestino { get; set; }
        public string dsDetalleDireccionDestino { get; set; }
        public string dsServicio { get; set; }
        public string dsUbicacionRiesgo { get; set; }
        public string dsTipoAsistencia { get; set; }
        public string dsRamo { get; set; }
        public string dsRiesgo { get; set; }
        public string dsMarcaVehiculo { get; set; }
        public string dsClaseVehiculo { get; set; }
        public string dsApliacionCreacion { get; set; }
        public string cdClienteAplicacionMovil { get; set; }
        public string dsEstadoServicio { get; set; }
        public string cdTipoDano { get; set; }
        public string dsCategoria { get; set; }
        public object marca { get; set; }
        public object dniUsuarioCreador { get; set; }
        public string distancia { get; set; }
        public object dsNombreProveedor { get; set; }
        public object dsNombreTecnico { get; set; }
        public object dsTiempoEnSitio { get; set; }
        public object conceptos { get; set; }
    }
}
