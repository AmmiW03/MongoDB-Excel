using MongoDB.Bson;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.Processing;
using System;
using System.Collections.Generic;
using System.Linq;
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
                String filepath = @"D:\Certificacion07-1\Excel\R01-Colaboradores.xlsx";

                FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);

                IWorkbook workbook = new XSSFWorkbook(fs);

                ISheet sheet = workbook.GetSheetAt(0);

                DataFormatter dataf = new DataFormatter();

                if (sheet == null) 
                    return;

                DateTime fecha = new DateTime(1, 1, 1, 0, 0, 0);
                DateTime fecha2 = new DateTime(1900, 1, 1, 0, 0, 0);

                IRow headRow = sheet.GetRow(5);
                for(int i = 6; i < sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    mdlEmplea.mdl_Emplea employee = new mdlEmplea.mdl_Emplea();
                    mdlEmplea.mdl_Beneficiario benefe = new mdlEmplea.mdl_Beneficiario();
                    for (int j = 0; i > row.Cells.Count; j++)
                    {
                        switch (dataf.FormatCellValue(headRow.GetCell(j)))
                        {
                            //Modulo 1

                            case "No."://Interpretación
                                employee.iEm_numero = int.Parse(dataf.FormatCellValue(row.GetCell(j)));//bd
                                break;

                            case "Nombre":
                                employee.sEm_nombre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Compañia":
                                employee.iEm_cia = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Fecha Ingreso":
                                employee.dtEm_fechai = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "Fecha Antiguedad":
                                employee.dtEm_fechant = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "Fecha Baja":
                                employee.dtEm_fechab = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "U.Camb Sal":
                                employee.dtEm_fechcam = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "Fecha Planta":
                                employee.dtEm_fecplan = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "Fecha U.Cont.":
                                employee.dtEm_feculco = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "E. Civ":
                                employee.sEm_estciv = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "RFC":
                                employee.sEm_rfc = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "No.Alfil IMSS":
                                employee.sEm_imss = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Gpo. IMSS":
                                employee.sEm_gruimss = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tip Col":
                                employee.sEm_tipoemp = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tip Nom":
                                employee.sEm_tiponom = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Puesto":
                                employee.iEm_puesto = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "No, Div":
                                employee.iEm_divisio = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Centro de Costo":
                                employee.iEm_depto = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "C. Pago":
                                employee.iEm_cpago = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Turno":
                                employee.sEm_turno = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "S.Diario":
                                employee.fEm_saldia = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.Propor":
                                employee.fEm_salprop = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.Prom.Prop":
                                employee.fEm_salppro = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.Prom":
                                employee.fEm_salprom = row.GetCell(j).NumericCellValue;
                                break;

                            case "Salario":
                                employee.fEm_salario = row.GetCell(j).NumericCellValue;
                                break;

                            case "T. Sal":
                                employee.sEm_tiposal = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "S.D.I":
                                employee.fEm_salinte = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.D.I. Var":
                                employee.fEm_sdivar = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.D.I. Ant":
                                employee.fEm_asalint = row.GetCell(j).NumericCellValue;
                                break;

                            case "S.D.I. Var Ant":
                                employee.fEm_avarant = row.GetCell(j).NumericCellValue;
                                break;

                            case "Sal. Ant":
                                employee.fEm_cambios = row.GetCell(j).NumericCellValue;
                                break;

                            case "F.Nac":
                                employee.dtEm_fechnac = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "Z Ec.":
                                employee.iEm_ubzona = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Suc":
                                employee.iEm_sucursa = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "M.O":
                                employee.sEm_manobra = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "T. San":
                                employee.sEm_tiposan = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tipo Cont":
                                employee.sEm_contra = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Nivel":
                                employee.iEm_nivel = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Reing":
                                employee.sEm_reingre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tab":
                                employee.iEm_tabula = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Sup":
                                employee.iEm_super = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "S. Garantia":
                                employee.fEm_salgara = row.GetCell(j).NumericCellValue;
                                break;

                            case "CURP":
                                employee.sEm_curp = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cel":
                                employee.sEm_celula = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Gpo":
                                employee.sEm_grupo = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Sub":
                                employee.sEm_subgrp = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "S.Tabulado":
                                employee.fEm_saltab = row.GetCell(j).NumericCellValue;
                                break;

                            case "Incentivo":
                                employee.fEm_incenti = row.GetCell(j).NumericCellValue;
                                break;

                            case "Dia Eco":
                                employee.iEm_diaeco = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "T. Col Ant":
                                employee.sEm_tempant = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "T. Nom Ant":
                                employee.sEm_tnomant = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "F. Camb. T.Col":
                                employee.dtEm_fnewnom = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "F.Matrimonio":
                                employee.dtEm_fecmatr = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;


                            //Modulo 2
                            case "C. ISPT":
                                employee.sEm_cispt = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Ajus Anu.":
                                employee.sEm_cajuste = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C. IMSS":
                                employee.sEm_cimss = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Pag C.S.":
                                employee.sEm_cfisica = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            //    Pendiente
                            //case "C.Rep PTU":
                            //    employee = ;
                            //    break;

                            case "C. Aguin":
                                employee.sEm_caguina = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "D. Aguin":
                                employee.iEm_cdias = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "C.V. Desp":
                                employee.sEm_valesde = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C.V. Com":
                                employee.sEm_valesco = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C.F. Aho.":
                                employee.sEm_cfahorr = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C.P. Asis.":
                                employee.iEm_asisten = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Imp Rec":
                                employee.sEm_irecibo = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Imp Cheq.":
                                employee.sEm_iforma = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Abon Ban":
                                employee.sEm_abonar = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "T. Ban":
                                employee.sEm_banco = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Suc Ban":
                                employee.fEm_sucurba = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Plaza Ban.E":
                                employee.sEm_plaza = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tipo Cta":
                                employee.sEm_tipocta = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cuenta Banco":
                                employee.sEm_cuenta = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C. INFO":
                                employee.sEm_cinfona = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cuenta. INFONAVIT":
                                employee.sEm_infocre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "% Dcto. Cred.IFONAVIT":
                                employee.fEm_infopor = row.GetCell(j).NumericCellValue;
                                break;

                            case "F.I.C. INFONAVIT":
                                employee.dtEm_fcreinf = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "T.Des. INFONAVIT":
                                employee.sEm_tipoinf = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "% Pen Alim.":
                                employee.fEm_penspor = row.GetCell(j).NumericCellValue;
                                break;

                            case "Importe Pen Alim":
                                employee.fEm_pensimp = row.GetCell(j).NumericCellValue;
                                break;

                            case "P.V. Aut":
                                employee.iEm_porprim = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            //Pendiente
                            //case "Ind. Vac":
                            //    employee.Add("");
                            //    break;

                            case "P.Re Vac":
                                employee.fEm_pereva = row.GetCell(j).NumericCellValue;
                                break;

                            case "Ini P.Vac":
                                employee.fEm_inpeva = row.GetCell(j).NumericCellValue;
                                break;

                            case "F.Ini.Vac.":
                                employee.dtEm_fechaiv = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "F.Reg.Vac.":
                                employee.dtEm_retvac = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue; ;
                                break;

                            case "SDI para 25 SMDF art 33 del SS":
                                employee.fEm_sdia29 = row.GetCell(j).NumericCellValue;
                                break;

                            case "SDI para 15 SDMF art 33 del SS":
                                employee.fEm_sdib29 = row.GetCell(j).NumericCellValue;
                                break;

                            case "% o cant. a pagar anticipo":
                                employee.fEm_anticip = row.GetCell(j).NumericCellValue;
                                break;

                            case "Cant. de Anti.Sem":
                                employee.fEm_antisem = row.GetCell(j).NumericCellValue;
                                break;

                            case "% Bono":
                                employee.fEm_porbono = row.GetCell(j).NumericCellValue;
                                break;

                            case "Factor Propor":
                                employee.fEm_propor = row.GetCell(j).NumericCellValue;
                                break;

                            case "Fac.C. Sal.Diario":
                                employee.fEm_minimom = row.GetCell(j).NumericCellValue;
                                break;

                            case "Reloj Chec.":
                                employee.sEm_reloj = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Act":
                                employee.sEm_activi = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cuenta Contable":
                                employee.sEm_ctacont = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "% Dcto. por Mant.":
                                employee.fEm_infoman = row.GetCell(j).NumericCellValue;
                                break;

                            case "Asimil. Salario":
                                employee.sEm_asimila = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "UMF":
                                employee.fEm_imssumf = row.GetCell(j).NumericCellValue;
                                break;

                            case "e-Mail e-Mail":
                                employee.sEm_email = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            //Modulo 3

                            case "Tel.  ":
                                employee.sRh_telefo = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Escolaridad":
                                employee.sRh_escolar = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cd. Nac.":
                                employee.sRh_nciudad = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Edo. Nac.":
                                employee.sRh_nestado = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Calle":
                                employee.sRh_dcalle = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Colonia":
                                employee.sRh_dcolon = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Ciudad":
                                employee.sRh_destado = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Estado":
                                employee.sRh_destado = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Municipio":
                                employee.sRh_dmunici = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "C.P.":
                                employee.sRh_dcp = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Nombre del padre":
                                employee.sRh_npadre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Nombre del madre":
                                employee.sRh_npadre = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "P Fin":
                                employee.iRh_fpadre = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "M Fin":
                                employee.iRh_fmadre = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Nacionalidad":
                                employee.sRh_nacion = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "1er. Asegurado GMM":
                                employee.sRh_gmmaseg = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "F.Nac. Asegurado":
                                employee.dtRh_gmmfnac = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue;
                                break;

                            case "Sexo":
                                if (employee.sEm_sexo == null)
                                {
                                    employee.sEm_sexo = dataf.FormatCellValue(row.GetCell(j));
                                }
                                else
                                {
                                    employee.sRh_gmmsexo = dataf.FormatCellValue(row.GetCell(j));
                                }
                                break;

                            case "Paren":
                                employee.sRh_gmmpare = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Nom. p/repor emergencia":
                                employee.fRh_gmmpor = row.GetCell(j).NumericCellValue;
                                break;

                            case "Primer Nombre p/emergencia":
                                employee.sRh_noavis1 = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tel.del emergente":
                                employee.sRh_teavis1 = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Parentesco":
                                if (employee.sRh_paavis1 == null)
                                {
                                    employee.sRh_paavis1 = dataf.FormatCellValue(row.GetCell(j));
                                }
                                else
                                {
                                    employee.sRh_paavis2 = dataf.FormatCellValue(row.GetCell(j));
                                }
                                break;

                            case "Segundo Nombre de emergencia":
                                employee.sRh_noavis1 = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Tel del emergente":
                                employee.sRh_teavis2 = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Fotografia Asoc.":
                                employee.sRh_picture = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Cve. P.GMM":
                                employee.sRh_gmmpcve = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Area Col":
                                employee.fRh_area = row.GetCell(j).NumericCellValue;
                                break;

                            case "Oficio":
                                employee.sRh_oficio = dataf.FormatCellValue(row.GetCell(j));
                                break;
                            
                            case "Estat":
                                benefe.fBe_estatur = row.GetCell(j).NumericCellValue;
                                break;

                            case "GMM":
                                employee.sRh_gmm = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Seg Vida":
                                employee.sRh_gmmsgv = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Suma  Aseg. GMM":
                                employee.fRh_gmmsuma = row.GetCell(j).NumericCellValue;
                                break;

                            case "Plan Seg.  Vida":
                                employee.iRh_plansv = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "P.A. S.V.":
                                employee.iRh_aplansv = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Prima Aseg.S.V.":
                                employee.fRh_psvsuma = int.Parse(dataf.FormatCellValue(row.GetCell(j)));
                                break;

                            case "Ubicación del Colaborador":
                                employee.sRh_ubicado = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Estat.":
                                employee.fRh_estatu = row.GetCell(j).NumericCellValue;
                                break;

                            case "Peso":
                                if (benefe.fBe_peso != null)
                                {
                                    benefe.fBe_peso = row.GetCell(j).NumericCellValue;
                                }
                                else
                                {
                                    employee.fRh_peso = row.GetCell(j).NumericCellValue;
                                }
                                break;

                            case "Talla Camisa":
                                employee.fRh_tallac = row.GetCell(j).NumericCellValue;
                                break;

                            case "Talla Pantalon":
                                employee.fRh_tallap = row.GetCell(j).NumericCellValue;
                                break;

                            case "Calzado":
                                employee.fRh_calzado = row.GetCell(j).NumericCellValue;
                                break;

                            case "Color Ojos":
                                employee.sRh_coloroj = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Color Cabello":
                                employee.sRh_colorca = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Color Piel":
                                employee.sRh_piel = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Señas Particulares":
                                employee.sRh_separt = dataf.FormatCellValue(row.GetCell(j));
                                break;

                            case "Estudio Soc-Eco":
                                employee.dtRh_soceco = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue;
                                break;

                            case "Alta Seg. Pub.":
                                employee.dtRh_segpub = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue;
                                break;

                            case "Examen Antidoping":
                                employee.dtRh_antidop = row.GetCell(j).DateCellValue == fecha ? fecha2 : row.GetCell(j).DateCellValue;
                                break;
                        }
                    }

                    BsonDocument document = new BsonDocument();
                    mdlMongoDB.SubirDatos("", employee);
                }
                
 
            }catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
