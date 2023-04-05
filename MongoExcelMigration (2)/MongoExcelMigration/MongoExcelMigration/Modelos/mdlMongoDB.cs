using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB;
using MongoDB.Driver;
using MongoDB.Bson;
using MongoDB.Driver.Linq;

namespace MongoExcelMigration.Modelos
{
    public static class mdlMongoDB
    {
        public static void SubirDatos(String sDBNombre,mdlEmplea.mdl_Emplea oDocumento)
        {
            try
            {
                MongoClient oClient = new MongoClient("mongodb://localhost:27017");

                var oDatabase = oClient.GetDatabase("Prueba");

                var coleccion = oDatabase.GetCollection<mdlEmplea.mdl_Emplea>("datos");
                //coleccion.InsertOne(oDocumento);
                coleccion.InsertOne(oDocumento);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
       
    }
}
