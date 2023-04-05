using MongoExcelMigration.Modelos;


Task tExcel = Task.Run(() => mdlMetodos.ReadExcel());

Console.WriteLine("\nCargando información...\n");

await tExcel;

if (tExcel.IsCompletedSuccessfully)
{
    Console.Clear();
    Console.WriteLine("\nDatos insertados correctamente.\n");
}