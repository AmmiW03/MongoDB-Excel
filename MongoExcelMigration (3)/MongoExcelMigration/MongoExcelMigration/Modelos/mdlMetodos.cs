using MongoDB.Bson;
using NPOI.HPSF;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.Processing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MongoExcelMigration.Modelos
{
    public static class mdlMetodos
    {
        public static void ReadExcel()
        {
            try
            {
                Console.WriteLine("\n Inicio del proceso... \n");
                String filepath = @"D:\Certificacion07-1\Excel\R01-Colaboradores.xlsx";

                FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);

                IWorkbook workbook = new XSSFWorkbook(fs);

                ISheet sheet = workbook.GetSheetAt(0);

                DataFormatter dataf = new DataFormatter();

                if (sheet == null)
                    return;

                DateTime fecha = new DateTime(1900, 1, 1, 0, 0, 0);
    
                IRow headRow = sheet.GetRow(5);
                for (int i = 6; i < sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    mdlEmplea.mdl_Emplea empleado = new mdlEmplea.mdl_Emplea();
                    mdlEmplea.mdl_Beneficiario[] beneficiario = new mdlEmplea.mdl_Beneficiario[]
                    {
                        new mdlEmplea.mdl_Beneficiario()
                    };
                    for (int j = 0; j < row.Cells.Count; j++)
                    {

                        switch (dataf.FormatCellValue(headRow.GetCell(j)))
                        {
                            //Modulo 1

                            case "No."://Interpretación
                                empleado.iEm_numero = int.Parse(dataf.FormatCellValue(row.GetCell(j)));//bd
                                break;

                            case "Nombre  ":
                                if (dataf.FormatCellValue(row.GetCell(j)) != null)
                                {
                                    empleado.sEm_nombre = dataf.FormatCellValue(row.GetCell(j));
                                }
                                else
                                {
                                    empleado.sEm_nombre = "";
                                }
                                break;

                            case "Compañía  ":
                                empleado.iEm_cia = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Fecha Ingreso":
                                empleado.dtEm_fechai = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                break;

                            case "Fecha Antigüedad":
                                empleado.dtEm_fechant = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                break;

                            case "Fecha Baja":
                                try
                                {
                                    empleado.dtEm_fechab = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_fechab = fecha;
                                }
                                break;

                            case "U.Camb Sal  ":
                                try
                                {
                                    empleado.dtEm_fechcam = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_fechcam = fecha;
                                }
                                break;
                            case "Fecha Planta":
                                try
                                {
                                    empleado.dtEm_fecplan = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_fecplan = fecha;
                                }
                                break;

                            case "Fecha U.Cont. ":
                                try
                                {
                                    empleado.dtEm_feculco = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_feculco = fecha;
                                }
                                break;

                            case "E. Civ":
                                empleado.sEm_estciv = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "RFC  ":
                                empleado.sEm_rfc = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            //case "No.Alfil IMSS":
                            //    empleado.sEm_imss = dataf.FormatCellValue(row.GetCell(j));
                            //    break;

                            case "No.Alfil IMSS":
                                if (empleado.sEm_imss == null)
                                {
                                    empleado.sEm_imss = "";
                                }
                                empleado.sEm_imss = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Gpo. IMSS":
                                empleado.sEm_gruimss = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tip Col":
                                empleado.sEm_tipoemp = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tip Nom":
                                empleado.sEm_tiponom = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Puesto":
                                empleado.iEm_puesto = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "No, Div":
                                empleado.iEm_divisio = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Centro de Costo":
                                empleado.iEm_depto = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "C. Pago":
                                int cpagoAux;
                                if (int.TryParse(dataf.FormatCellValue(row.GetCell(j)), out cpagoAux))
                                    empleado.iEm_cpago = cpagoAux;
                                else
                                    empleado.iEm_cpago = cpagoAux;
                                break;

                            case "Turno":
                                empleado.sEm_turno = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "S.Diario":
                                empleado.fEm_saldia = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.Propor":
                                empleado.fEm_salprop = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.Prom.Prop":
                                empleado.fEm_salppro = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.Prom":
                                empleado.fEm_salprom = row.GetCell(j).NumericCellValue;
                                break;

                            case "Salario  ":
                                empleado.fEm_salario = row.GetCell(j).NumericCellValue;
                                break;

                            case "T. Sal ":
                                empleado.sEm_tiposal = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "S.D.I.":
                                empleado.fEm_salinte = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.D.I.Var":
                                empleado.fEm_sdivar = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.D.I.Ant  ":
                                empleado.fEm_asalint = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.D.I.Var Ant  ":
                                empleado.fEm_avarant = row.GetCell(j).NumericCellValue;
                                break;

                            case "Sal. Ant  ":
                                empleado.fEm_cambios = row.GetCell(j).NumericCellValue;
                                break;

                            case "F.Nac  ":
                                try
                                {
                                    empleado.dtEm_fechnac = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_fechnac = fecha;
                                }
                                break;
                                break;

                            case "Z Ec.":
                                empleado.iEm_ubzona = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Suc":
                                empleado.iEm_sucursa = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "M.O":
                                empleado.sEm_manobra = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "T. San":
                                empleado.sEm_tiposan = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tipo Cont":
                                empleado.sEm_contra = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Nivel":
                                empleado.iEm_nivel = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Reing  ":
                                empleado.sEm_reingre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tab":
                                empleado.iEm_tabula = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Super":
                                empleado.iEm_super = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "S.Garantía  ":
                                empleado.fEm_salgara = row.GetCell(j).NumericCellValue;
                                break;

                            case "CURP  ":
                                empleado.sEm_curp = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cel  ":
                                empleado.sEm_celula = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Gpo  ":
                                empleado.sEm_grupo = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Sub  ":
                                empleado.sEm_subgrp = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "S.Tabulado  ":
                                empleado.fEm_saltab = row.GetCell(j).NumericCellValue;
                                break;

                            case "Incentivo  ":
                                empleado.fEm_incenti = row.GetCell(j).NumericCellValue;
                                break;

                            case "Día Eco":
                                empleado.iEm_diaeco = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "T. Col Ant":
                                empleado.sEm_tempant = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "T. Nom Ant":
                                empleado.sEm_tnomant = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "F. Camb. T.Col":
                                try
                                {
                                    empleado.dtEm_fnewnom = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_fnewnom = fecha;
                                }
                                break;

                            case "F.Matrimonio  ":
                                try
                                {
                                    empleado.dtEm_fecmatr = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_fecmatr = fecha;
                                }
                                break;

                            //Modulo 2
                            case "C. ISPT":
                                empleado.sEm_cispt = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Ajus Anu.":
                                empleado.sEm_cajuste = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C. IMSS":
                                empleado.sEm_cimss = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Pag C.S.":
                                empleado.sEm_cfisica = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            //    Pendiente
                            //case "C.Rep PTU":
                            //    empleado = ;
                            //    break;

                            case "C. Aguin":
                                empleado.sEm_caguina = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "D. Aguin":
                                empleado.iEm_cdias = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "C.V. Desp":
                                empleado.sEm_valesde = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C.V. Com":
                                empleado.sEm_valesco = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C.F. Aho.":
                                empleado.sEm_cfahorr = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C.P. Asis.":
                                int cpAsisAux;
                                if (int.TryParse(dataf.FormatCellValue(row.GetCell(j)), out cpAsisAux))
                                    empleado.iEm_asisten = cpAsisAux;
                                else
                                    empleado.iEm_asisten = cpAsisAux;
                                break;

                            case "Imp Rec":
                                empleado.sEm_irecibo = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Imp Cheq.":
                                empleado.sEm_iforma = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Abon Ban":
                                empleado.sEm_abonar = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "T. Ban":
                                empleado.sEm_banco = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Suc Ban":
                                empleado.fEm_sucurba = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Plaza Ban.E":
                                empleado.sEm_plaza = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tipo Cta":
                                empleado.sEm_tipocta = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cuenta Banco  ":
                                empleado.sEm_cuenta = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C. INFO":
                                empleado.sEm_cinfona = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cuenta. INFONAVIT":
                                empleado.sEm_infocre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "% Dcto. Cred.IFONAVIT":
                                empleado.fEm_infopor = row.GetCell(j).NumericCellValue;
                                break;

                            case "F.I.C. INFONAVIT":
                                try
                                {
                                    empleado.dtEm_fcreinf = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_fcreinf = fecha;
                                }
                                break;

                            case "T.Des. INFONAVIT":
                                empleado.sEm_tipoinf = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "% Pen Alim.":
                                empleado.fEm_penspor = row.GetCell(j).NumericCellValue;
                                break;

                            case "Importe Pen Alim":
                                empleado.fEm_pensimp = row.GetCell(j).NumericCellValue;
                                break;

                            case "P.V. Aut":
                                int pvAutAux;
                                if (int.TryParse(dataf.FormatCellValue(row.GetCell(j)), out pvAutAux))
                                    empleado.iEm_porprim = pvAutAux;
                                else
                                    empleado.iEm_porprim = pvAutAux;
                                break;

                            //Pendiente
                            //case "Ind. Vac":
                            //    empleado.Add("");
                            //    break;

                            case "P.Re Vac":
                                empleado.fEm_pereva = row.GetCell(j).NumericCellValue;
                                break;

                            case "Ini P.Vac":
                                empleado.fEm_inpeva = row.GetCell(j).NumericCellValue;
                                break;

                            case "F.Ini.Vac.  ":
                                try
                                {
                                    empleado.dtEm_fechaiv = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_fechaiv = fecha;
                                }
                                break;

                            case "F.Reg.Vac.  ":
                                try
                                {
                                    empleado.dtEm_retvac = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtEm_retvac = fecha;
                                }
                                break;

                            case "SDI para 25 SMDF art 33 del SS":
                                empleado.fEm_sdia29 = row.GetCell(j).NumericCellValue;
                                break;

                            case "SDI para 15 SDMF art 33 del SS":
                                empleado.fEm_sdib29 = row.GetCell(j).NumericCellValue;
                                break;

                            case "% o cant. a pagar anticipo":
                                empleado.fEm_anticip = row.GetCell(j).NumericCellValue;
                                break;

                            case "Cant. de Anti.Sem":
                                empleado.fEm_antisem = row.GetCell(j).NumericCellValue;
                                break;

                            case "% Bono ":
                                empleado.fEm_porbono = row.GetCell(j).NumericCellValue;
                                break;

                            case "Factor Propor  ":
                                empleado.fEm_propor = row.GetCell(j).NumericCellValue;
                                break;

                            case "Fac.C. Sal.Diario":
                                empleado.fEm_minimom = row.GetCell(j).NumericCellValue;
                                break;

                            case "Reloj Chec.":
                                empleado.sEm_reloj = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Act  ":
                                empleado.sEm_activi = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cuenta Contable":
                                empleado.sEm_ctacont = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "% Dcto. por Mant.":
                                empleado.fEm_infoman = row.GetCell(j).NumericCellValue;
                                break;

                            case "Asimil. Salario":
                                empleado.sEm_asimila = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "UMF  ":
                                empleado.fEm_imssumf = row.GetCell(j).NumericCellValue;
                                break;

                            case "e-Mail e-Mail":
                                empleado.sEm_email = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            //Modulo 3

                            case "Tel.  ":
                                empleado.sRh_telefo = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Escolaridad  ":
                                empleado.sRh_escolar = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cd. Nac.":
                                empleado.sRh_nciudad = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Edo. Nac.  ":
                                empleado.sRh_nestado = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Calle  ":
                                empleado.sRh_dcalle = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Colonia  ":
                                empleado.sRh_dcolon = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Ciudad  ":
                                empleado.sRh_destado = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Estado  ":
                                empleado.sRh_destado = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Municipio  ":
                                empleado.sRh_dmunici = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C.P.  ":
                                empleado.sRh_dcp = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Nombre del padre":
                                empleado.sRh_npadre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Nombre del madre":
                                empleado.sRh_npadre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "P Fin":
                                int pFinAux;
                                if (int.TryParse(dataf.FormatCellValue(row.GetCell(j)), out pFinAux))
                                    empleado.iRh_fpadre = pFinAux;
                                else
                                    empleado.iRh_fpadre = pFinAux;
                                break;

                            case "M Fin":
                                int MFinAux;
                                if (int.TryParse(dataf.FormatCellValue(row.GetCell(j)), out pFinAux))
                                    empleado.iRh_fmadre = pFinAux;
                                else
                                    empleado.iRh_fmadre = pFinAux;
                                break;

                            case "Nacionalidad  ":
                                empleado.sRh_nacion = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "1er. Asegurado GMM  ":
                                empleado.sRh_gmmaseg = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "F.Nac. Asegurado":
                                try
                                {
                                    empleado.dtRh_gmmfnac = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtRh_gmmfnac = fecha;
                                }
                                break;

                            case "Sexo":
                                if (empleado.sEm_sexo == null)
                                {
                                    empleado.sEm_sexo = dataf.FormatCellValue(row.GetCell(j));
                                }
                                else
                                {
                                    empleado.sRh_gmmsexo = dataf.FormatCellValue(row.GetCell(j));
                                }
                                break;

                            case "Paren  ":
                                empleado.sRh_gmmpare = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Nom. p/repor emergencia":
                                empleado.fRh_gmmpor = row.GetCell(j).NumericCellValue;
                                break;

                            case "Primer Nombre p/emergencia":
                                empleado.sRh_noavis1 = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tel.del emergente":
                                empleado.sRh_teavis1 = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Parentesco":
                                if (empleado.sRh_paavis1 == null)
                                {
                                    empleado.sRh_paavis1 = dataf.FormatCellValue(row.GetCell(j));
                                }
                                else
                                {
                                    empleado.sRh_paavis2 = dataf.FormatCellValue(row.GetCell(j));
                                }
                                break;

                            case "Segundo Nombre de emergencia":
                                empleado.sRh_noavis1 = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tel del emergente":
                                empleado.sRh_teavis2 = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Fotografia Asoc.":
                                empleado.sRh_picture = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cve. P.GMM":
                                empleado.sRh_gmmpcve = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Area Col":
                                empleado.fRh_area = row.GetCell(j).NumericCellValue;
                                break;

                            case "Oficio  ":
                                empleado.sRh_oficio = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Estat  ":
                                beneficiario[0].fBe_estatur = row.GetCell(j).NumericCellValue;
                                break;

                            case "GMM  ":
                                empleado.sRh_gmm = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Seg Vida":
                                empleado.sRh_gmmsgv = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Suma  Aseg. GMM":
                                empleado.fRh_gmmsuma = row.GetCell(j).NumericCellValue;
                                break;

                            case "Plan Seg.  Vida":
                                int plansvAux;
                                if (int.TryParse(dataf.FormatCellValue(row.GetCell(j)), out plansvAux))
                                    empleado.iRh_plansv = plansvAux;
                                else
                                    empleado.iRh_plansv = plansvAux;
                                break;

                            case "P.A. S.V.":
                                int aplansvAux;
                                if (int.TryParse(dataf.FormatCellValue(row.GetCell(j)), out aplansvAux))
                                    empleado.iRh_aplansv = aplansvAux;
                                else
                                    empleado.iRh_aplansv = aplansvAux;
                                break;

                            case "Prima Aseg.S.V.":
                                int psvSumAux;
                                if (int.TryParse(dataf.FormatCellValue(row.GetCell(j)), out psvSumAux))
                                    empleado.fRh_psvsuma = psvSumAux;
                                else
                                    empleado.fRh_psvsuma = psvSumAux;
                                break;

                            case "Ubicación del Colaborador":
                                empleado.sRh_ubicado = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Estat.  ":
                                empleado.fRh_estatu = row.GetCell(j).NumericCellValue;
                                break;

                            case "Peso":
                                if (beneficiario[0].fBe_peso != null)
                                {
                                    beneficiario[0].fBe_peso = row.GetCell(j).NumericCellValue;
                                }
                                else
                                {
                                    empleado.fRh_peso = row.GetCell(j).NumericCellValue;
                                }
                                break;

                            case "Talla Camisa":
                                empleado.fRh_tallac = row.GetCell(j).NumericCellValue;
                                break;

                            case "Talla Pantalon":
                                empleado.fRh_tallap = row.GetCell(j).NumericCellValue;
                                break;

                            case "Calzado  ":
                                empleado.fRh_calzado = row.GetCell(j).NumericCellValue;
                                break;

                            case "Color Ojos":
                                empleado.sRh_coloroj = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Color Cabello":
                                empleado.sRh_colorca = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Color Piel":
                                empleado.sRh_piel = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Señas Particulares":
                                empleado.sRh_separt = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Estudio Soc-Eco":
                                try
                                {
                                    empleado.dtRh_soceco = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtRh_soceco = fecha;
                                }
                                break;

                            case "Alta Seg. Pub.":
                                try
                                {
                                    empleado.dtRh_segpub = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtRh_segpub = fecha;
                                }
                                break;

                            case "Examen Antidoping":
                                try
                                {
                                    empleado.dtRh_antidop = DateTime.ParseExact(row.GetCell(j).StringCellValue, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                }
                                catch
                                {
                                    empleado.dtRh_antidop = fecha;
                                }
                                break;
                        }
                        //Console.WriteLine("Finalizacion de una columna");
                    }
                    empleado.aBeneficiarios = beneficiario;
                    PropertyInfo[] propiedades = typeof(mdlEmplea.mdl_Emplea).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    foreach (PropertyInfo propiedad in propiedades)
                    {
                        if (propiedad.GetValue(empleado) == null)
                        {
                            switch (propiedad.PropertyType)
                            {
                                case Type t when t == typeof(string):
                                    propiedad.SetValue(empleado, string.Empty);
                                    break;
                                case Type t when t == typeof(int):
                                    propiedad.SetValue(empleado, 0);
                                    break;
                                case Type t when t == typeof(double):
                                    propiedad.SetValue(empleado, 0.0);
                                    break;
                                case Type t when t == typeof(DateTime):
                                    propiedad.SetValue(empleado, new DateTime(1900, 1, 1, 0, 0, 0));
                                    break;
                            }
                        }

                        if (propiedad.PropertyType == typeof(DateTime))
                        {
                            if ((DateTime)propiedad.GetValue(empleado) == DateTime.MinValue)
                            {
                                propiedad.SetValue(empleado, DateTime.ParseExact("01/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture));
                            }
                        }
                    }

                    mdlEmplea.mdl_Conceptos[] concepto = new mdlEmplea.mdl_Conceptos[]
                                    {
                                        new mdlEmplea.mdl_Conceptos()
                                        {
                                            iPd_periodo = 0,
                                        }
                                    };
                    empleado.aConceptos = concepto;
                    PropertyInfo[] propertieA = typeof(mdlEmplea.mdl_Conceptos).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    foreach (PropertyInfo propiedad in propertieA)
                    {
                        if (propiedad.GetValue(concepto[0]) == null)
                        {
                            switch (propiedad.PropertyType)
                            {
                                case Type t when t == typeof(string):
                                    propiedad.SetValue(concepto[0], string.Empty);
                                    break;
                                case Type t when t == typeof(int):
                                    propiedad.SetValue(concepto[0], 0);
                                    break;
                                case Type t when t == typeof(double):
                                    propiedad.SetValue(concepto[0], 0.0);
                                    break;
                                case Type t when t == typeof(long):
                                    propiedad.SetValue(concepto[0], 0);
                                    break;
                                case Type t when t == typeof(DateTime):
                                    propiedad.SetValue(concepto[0], DateTime.ParseExact("01/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture));
                                    break;
                            }
                        }
                    }

                    mdlEmplea.mdl_Vacaciones[] vacaciones = new mdlEmplea.mdl_Vacaciones[]
                                    {
                                        new mdlEmplea.mdl_Vacaciones()
                                        {
                                            iEm_vdepto = 0,
                                        }
                                    };
                    empleado.aVacaciones = vacaciones;
                    PropertyInfo[] propertieB = typeof(mdlEmplea.mdl_Vacaciones).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    foreach (PropertyInfo propiedad in propertieB)
                    {
                        if (propiedad.GetValue(vacaciones[0]) == null)
                        {
                            switch (propiedad.PropertyType)
                            {
                                case Type t when t == typeof(string):
                                    propiedad.SetValue(vacaciones[0], string.Empty);
                                    break;
                                case Type t when t == typeof(int):
                                    propiedad.SetValue(concepto[0], 0);
                                    break;
                                case Type t when t == typeof(double):
                                    propiedad.SetValue(concepto[0], 0.0);
                                    break;
                                case Type t when t == typeof(long):
                                    propiedad.SetValue(concepto[0], 0);
                                    break;
                                case Type t when t == typeof(DateTime):
                                    propiedad.SetValue(concepto[0], DateTime.ParseExact("01/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture));
                                    break;
                            }
                        }
                    }

                    mdlEmplea.mdl_Salario[] salario = new mdlEmplea.mdl_Salario[]
                                    {
                                        new mdlEmplea.mdl_Salario()
                                        {
                                             iEm_topes = 0,
                                        }
                                    };
                    empleado.aIncremento = salario;
                    PropertyInfo[] propertieC = typeof(mdlEmplea.mdl_Salario).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    foreach (PropertyInfo propiedad in propertieC)
                    {
                        if (propiedad.GetValue(salario[0]) == null)
                        {
                            switch (propiedad.PropertyType)
                            {
                                case Type t when t == typeof(string):
                                    propiedad.SetValue(salario[0], string.Empty);
                                    break;
                                case Type t when t == typeof(int):
                                    propiedad.SetValue(salario[0], 0);
                                    break;
                                case Type t when t == typeof(double):
                                    propiedad.SetValue(salario[0], 0.0);
                                    break;
                                case Type t when t == typeof(long):
                                    propiedad.SetValue(salario[0], 0);
                                    break;
                                case Type t when t == typeof(DateTime):
                                    propiedad.SetValue(salario[0], DateTime.ParseExact("01/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture));
                                    break;
                            }
                        }
                    }

                    mdlEmplea.mdl_Historico[] historico = new mdlEmplea.mdl_Historico[]
                                    {
                                        new mdlEmplea.mdl_Historico()
                                        {
                                             sEm_tipomov = "",
                                        }
                                    };
                    empleado.aHistorico = historico;
                    PropertyInfo[] propertieD = typeof(mdlEmplea.mdl_Historico).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    foreach (PropertyInfo propiedad in propertieD)
                    {
                        if (propiedad.GetValue(historico[0]) == null)
                        {
                            switch (propiedad.PropertyType)
                            {
                                case Type t when t == typeof(string):
                                    propiedad.SetValue(historico[0], string.Empty);
                                    break;
                                case Type t when t == typeof(int):
                                    propiedad.SetValue(historico[0], 0);
                                    break;
                                case Type t when t == typeof(double):
                                    propiedad.SetValue(historico[0], 0.0);
                                    break;
                                case Type t when t == typeof(long):
                                    propiedad.SetValue(historico[0], 0);
                                    break;
                                case Type t when t == typeof(DateTime):
                                    propiedad.SetValue(historico[0], DateTime.ParseExact("01/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture));
                                    break;
                            }
                        }
                    }

                    mdlEmplea.mdl_Reingreso[] reingreso = new mdlEmplea.mdl_Reingreso[]
                                    {
                                        new mdlEmplea.mdl_Reingreso()
                                        {
                                             sEm_rcausa = "",
                                        }
                                    };
                    empleado.aReingreso = reingreso;
                    PropertyInfo[] propertieE = typeof(mdlEmplea.mdl_Reingreso).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    foreach (PropertyInfo propiedad in propertieE)
                    {
                        if (propiedad.GetValue(reingreso[0]) == null)
                        {
                            switch (propiedad.PropertyType)
                            {
                                case Type t when t == typeof(string):
                                    propiedad.SetValue(reingreso[0], string.Empty);
                                    break;
                                case Type t when t == typeof(int):
                                    propiedad.SetValue(reingreso[0], 0);
                                    break;
                                case Type t when t == typeof(double):
                                    propiedad.SetValue(reingreso[0], 0.0);
                                    break;
                                case Type t when t == typeof(long):
                                    propiedad.SetValue(reingreso[0], 0);
                                    break;
                                case Type t when t == typeof(DateTime):
                                    propiedad.SetValue(reingreso[0], DateTime.ParseExact("01/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture));
                                    break;
                            }
                        }
                    }

                    mdlEmplea.mdl_Baja[] baja = new mdlEmplea.mdl_Baja[]
                                    {
                                        new mdlEmplea.mdl_Baja()
                                        {
                                             sEm_bcausa = "",
                                        }
                                    };
                    empleado.aBajas = baja;
                    PropertyInfo[] propertieF = typeof(mdlEmplea.mdl_Baja).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    foreach (PropertyInfo propiedad in propertieF)
                    {
                        if (propiedad.GetValue(baja[0]) == null)
                        {
                            switch (propiedad.PropertyType)
                            {
                                case Type t when t == typeof(string):
                                    propiedad.SetValue(baja[0], string.Empty);
                                    break;
                                case Type t when t == typeof(int):
                                    propiedad.SetValue(baja[0], 0);
                                    break;
                                case Type t when t == typeof(double):
                                    propiedad.SetValue(baja[0], 0.0);
                                    break;
                                case Type t when t == typeof(long):
                                    propiedad.SetValue(baja[0], 0);
                                    break;
                                case Type t when t == typeof(DateTime):
                                    propiedad.SetValue(baja[0], DateTime.ParseExact("01/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture));
                                    break;
                            }
                        }
                    }

                    BsonDocument document = new BsonDocument();
                        mdlMongoDB.SubirDatos(empleado);
                }

                    Console.WriteLine("\n Proceso finalizado...");
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}