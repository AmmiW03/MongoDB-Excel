using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MongoExcelMigration.Modelos
{
    public class mdlEmplea
    {
        public class mdl_Emplea
        {
            [BsonId]
            [BsonRepresentation(BsonType.ObjectId)]
            public string sId { get; set; }

            #region nemplea
            [BsonElement("em_numant")]
            public int iEm_numant { get; set; }
            [BsonElement("em_numero")]
            public int iEm_numero { get; set; }

            [BsonElement("em_cia")]
            public int iEm_cia { get; set; }

            [BsonElement("em_nombre")]
            public string sEm_nombre { get; set; }

            [BsonElement("em_fechai")]
            public DateTime dtEm_fechai { get; set; }

            [BsonElement("em_imfecha")]
            public DateTime dtEm_imfecha { get; set; }

            [BsonElement("em_fechab")]
            public DateTime dtEm_fechab { get; set; }

            [BsonElement("em_fecnac")]
            public DateTime dtEm_fecnac { get; set; }

            [BsonElement("em_fechcam")]
            public DateTime dtEm_fechcam { get; set; }

            [BsonElement("em_fecplan")]
            public DateTime dtEm_fecplan { get; set; }

            [BsonElement("em_feculco")]
            public DateTime dtEm_feculco { get; set; }

            [BsonElement("em_estciv")]
            public string sEm_estciv { get; set; }

            [BsonElement("em_rfc")]
            public string sEm_rfc { get; set; }

            [BsonElement("em_sar")]
            public string sEm_sar { get; set; }

            [BsonElement("em_imss")]
            public string sEm_imss { get; set; }

            [BsonElement("em_gruimss")]
            public string sEm_gruimss { get; set; }

            [BsonElement("em_tipoemp")]
            public string sEm_tipoemp { get; set; }

            [BsonElement("em_tiponom")]
            public string sEm_tiponom { get; set; }

            [BsonElement("em_puesto")]
            public int iEm_puesto { get; set; }

            [BsonElement("em_divisio")]
            public int iEm_divisio { get; set; }

            [BsonElement("em_depto")]
            public int iEm_depto { get; set; }

            [BsonElement("em_turno")]
            public string sEm_turno { get; set; }

            [BsonElement("em_salario")]
            public double fEm_salario { get; set; }

            [BsonElement("em_saldia")]
            public double fEm_saldia { get; set; }

            [BsonElement("em_salprop")]
            public double fEm_salprop { get; set; }

            [BsonElement("em_salprom")]
            public double fEm_salprom { get; set; }

            [BsonElement("em_salppro")]
            public double fEm_salppro { get; set; }

            [BsonElement("em_tiposal")]
            public string sEm_tiposal { get; set; }

            [BsonElement("em_salinte")]
            public double fEm_salinte { get; set; }

            [BsonElement("em_sdivar")]
            public double fEm_sdivar { get; set; }

            [BsonElement("em_asalint")]
            public double fEm_asalint { get; set; }

            [BsonElement("em_avarant")]
            public double fEm_avarant { get; set; }

            [BsonElement("em_cambios")]
            public double fEm_cambios { get; set; }

            [BsonElement("em_sexo")]
            public string sEm_sexo { get; set; }

            [BsonElement("em_fechnac")]
            public DateTime dtEm_fechnac { get; set; }

            [BsonElement("em_ubzona")]
            public int iEm_ubzona { get; set; }

            [BsonElement("em_sucursa")]
            public int iEm_sucursa { get; set; }

            [BsonElement("em_manobra")]
            public string sEm_manobra { get; set; }

            [BsonElement("em_tiposan")]
            public string sEm_tiposan { get; set; }

            [BsonElement("em_contra")]
            public string sEm_contra { get; set; }

            [BsonElement("em_tiposnc")]
            public string sEm_tiposnc { get; set; }

            [BsonElement("em_grupoim")]
            public string sEm_grupoim { get; set; }

            [BsonElement("em_nivel")]
            public int iEm_nivel { get; set; }

            [BsonElement("em_sdifij1")]
            public double fEm_sdifij1 { get; set; }

            [BsonElement("em_sdifij2")]
            public double fEm_sdifij2 { get; set; }

            [BsonElement("em_sdifij3")]
            public double fEm_sdifij3 { get; set; }

            [BsonElement("em_sdifij4")]
            public double fEm_sdifij4 { get; set; }

            [BsonElement("em_sdifij5")]
            public double fEm_sdifij5 { get; set; }

            [BsonElement("em_sdifij6")]
            public double fEm_sdifij6 { get; set; }

            [BsonElement("em_sdibim1")]
            public double fEm_sdibim1 { get; set; }

            [BsonElement("em_sdibim2")]
            public double fEm_sdibim2 { get; set; }

            [BsonElement("em_sdibim3")]
            public double fEm_sdibim3 { get; set; }

            [BsonElement("em_sdibim4")]
            public double fEm_sdibim4 { get; set; }

            [BsonElement("em_sdibim5")]
            public double fEm_sdibim5 { get; set; }

            [BsonElement("em_sdibim6")]
            public double fEm_sdibim6 { get; set; }

            [BsonElement("em_status")]
            public string sEm_status { get; set; }

            [BsonElement("em_reingre")]
            public string sEm_reingre { get; set; }

            [BsonElement("em_tabula")]
            public int iEm_tabula { get; set; }

            [BsonElement("em_super")]
            public int iEm_super { get; set; }

            [BsonElement("em_area")]
            public int iEm_area { get; set; }

            [BsonElement("em_sdinfo1")]
            public double fEm_sdinfo1 { get; set; }

            [BsonElement("em_sdinfo2")]
            public double fEm_sdinfo2 { get; set; }

            [BsonElement("em_sdinfo3")]
            public double fEm_sdinfo3 { get; set; }

            [BsonElement("em_sdinfo4")]
            public double fEm_sdinfo4 { get; set; }

            [BsonElement("em_sdinfo5")]
            public double fEm_sdinfo5 { get; set; }

            [BsonElement("em_sdinfo6")]
            public double fEm_sdinfo6 { get; set; }

            [BsonElement("em_salgara")]
            public double fEm_salgara { get; set; }

            [BsonElement("em_curp")]
            public string sEm_curp { get; set; }

            [BsonElement("em_celula")]
            public string sEm_celula { get; set; }

            [BsonElement("em_grupo")]
            public string sEm_grupo { get; set; }

            [BsonElement("em_subgrp")]
            public string sEm_subgrp { get; set; }

            [BsonElement("m_saltab")]
            public double fEm_saltab { get; set; }

            [BsonElement("em_incenti")]
            public double fEm_incenti { get; set; }

            [BsonElement("em_diaeco")]
            public int iEm_diaeco { get; set; }

            [BsonElement("em_tempant")]
            public string sEm_tempant { get; set; }

            [BsonElement("em_tnomant")]
            public string sEm_tnomant { get; set; }

            [BsonElement("em_fnewemp")]
            public DateTime dtEm_fnewemp { get; set; }

            [BsonElement("em_fnewnom")]
            public DateTime dtEm_fnewnom { get; set; }

            [BsonElement("em_fecmatr")]
            public DateTime dtEm_fecmatr { get; set; }

            [BsonElement("em_fechori")]
            public DateTime dtEm_fechori { get; set; }

            [BsonElement("em_fecimss")]
            public DateTime dtEm_fecimss { get; set; }

            [BsonElement("em_salxhor")]
            public double fEm_salxhor { get; set; }

            [BsonElement("em_tipojor")]
            public string sEm_tipojor { get; set; }

            [BsonElement("em_asimila")]
            public string sEm_asimila { get; set; }

            [BsonElement("em_sede")]
            public string sEm_sede { get; set; }

            [BsonElement("em_lugar")]
            public string sEm_lugar { get; set; }

            [BsonElement("em_pasport")]
            public string sEm_pasport { get; set; }

            [BsonElement("em_fecpass")]
            public DateTime dtEm_fecpass { get; set; }

            [BsonElement("em_certifi")]
            public string sEm_certifi { get; set; }

            [BsonElement("em_ciudad")]
            public string sEm_ciudad { get; set; }

            [BsonElement("em_usuario")]
            public string sEm_usuario { get; set; }

            [BsonElement("em_autom")]
            public int iEm_autom { get; set; }

            [BsonElement("em_email")]
            public string sEm_email { get; set; }

            [BsonElement("em_fechant")]
            public DateTime dtEm_fechant { get; set; }

            [BsonElement("em_observa")]
            public string sEm_observa { get; set; }

            [BsonElement("em_horasxp")]
            public double fEm_horasxp { get; set; }

            [BsonElement("em_sneto")]
            public double fEm_sneto { get; set; }

            [BsonElement("em_diasvac")]
            public int iEm_diasvac { get; set; }

            [BsonElement("em_porprim")]
            public int iEm_porprim { get; set; }

            [BsonElement("em_tipoasi")]
            public int iEm_tipoasi { get; set; }

            [BsonElement("em_nomsat")]
            public string sEm_nomsat { get; set; }

            [BsonElement("em_cpsat")]
            public string sEm_cpsat { get; set; }
            #endregion

            #region nemppar
            [BsonElement("em_cispt")]
            public string sEm_cispt { get; set; }

            [BsonElement("em_cajuste")]
            public string sEm_cajuste { get; set; }

            [BsonElement("em_cimss")]
            public string sEm_cimss { get; set; }

            [BsonElement("em_cinfona")]
            public string sEm_cinfona { get; set; }

            [BsonElement("em_cparti")]
            public string sEm_cparti { get; set; }

            [BsonElement("em_caguina")]
            public string sEm_caguina { get; set; }

            [BsonElement("em_cdias")]
            public int iEm_cdias { get; set; }

            [BsonElement("em_cfisica")]
            public string sEm_cfisica { get; set; }

            [BsonElement("em_cconsta")]
            public string sEm_cconsta { get; set; }

            [BsonElement("em_cfahorr")]
            public string sEm_cfahorr { get; set; }

            [BsonElement("em_iforma")]
            public string sEm_iforma { get; set; }

            [BsonElement("em_cuenta")]
            public string sEm_cuenta { get; set; }

            [BsonElement("em_cuenta2")]
            public string sEm_cuenta2 { get; set; }

            [BsonElement("em_irecibo")]
            public string sEm_irecibo { get; set; }

            [BsonElement("em_pvacaci")]
            public string sEm_pvacaci { get; set; }

            [BsonElement("em_propor")]
            public double fEm_propor { get; set; }

            [BsonElement("em_minimom")]
            public double fEm_minimom { get; set; }

            [BsonElement("em_valesde")]
            public string sEm_valesde { get; set; }

            [BsonElement("em_intdes")]
            public string sEm_intdes { get; set; }

            [BsonElement("em_valesco")]
            public string sEm_valesco { get; set; }

            [BsonElement("em_banco")]
            public string sEm_banco { get; set; }

            [BsonElement("em_abonar")]
            public string sEm_abonar { get; set; }

            [BsonElement("em_reloj")]
            public string sEm_reloj { get; set; }

            [BsonElement("em_infocre")]
            public string sEm_infocre { get; set; }

            [BsonElement("em_infopor")]
            public double fEm_infopor { get; set; }

            [BsonElement("em_infoman")]
            public double fEm_infoman { get; set; }

            [BsonElement("em_sucurba")]
            public string fEm_sucurba { get; set; }

            [BsonElement("em_plaza")]
            public string sEm_plaza { get; set; }

            [BsonElement("em_plaza2")]
            public string sEm_plaza2 { get; set; }

            [BsonElement("em_porbco1")]
            public double fEm_porbco1 { get; set; }

            [BsonElement("em_porbco2")]
            public double fEm_porbco2 { get; set; }

            [BsonElement("em_penspor")]
            public double fEm_penspor { get; set; }

            [BsonElement("em_pensimp")]
            public double fEm_pensimp { get; set; }

            [BsonElement("em_pentipo")]
            public double fEm_pentipo { get; set; }

            [BsonElement("em_movfij")]
            public string sEm_movfij { get; set; }

            [BsonElement("em_envaca")]
            public string sEm_envaca { get; set; }

            [BsonElement("em_pereva")]
            public double fEm_pereva { get; set; }

            [BsonElement("em_inpeva")]
            public double fEm_inpeva { get; set; }

            [BsonElement("em_anticip")]
            public double fEm_anticip { get; set; }

            [BsonElement("em_sdia29")]
            public double fEm_sdia29 { get; set; }

            [BsonElement("em_sdib29")]
            public double fEm_sdib29 { get; set; }

            [BsonElement("em_activi")]
            public string sEm_activi { get; set; }

            [BsonElement("em_cpago")]
            public int iEm_cpago { get; set; }

            [BsonElement("em_penxinv")]
            public string sEm_penxinv { get; set; }

            [BsonElement("em_nocredi")]
            public string sEm_nocredi { get; set; }

            [BsonElement("em_ificha")]
            public string sEm_ificha { get; set; }

            [BsonElement("em_antisem")]
            public double fEm_antisem { get; set; }

            [BsonElement("em_facsm")]
            public double fEm_facsm { get; set; }

            [BsonElement("em_ctacont")]
            public string sEm_ctacont { get; set; }

            [BsonElement("em_fechaiv")]
            public DateTime dtEm_fechaiv { get; set; }

            [BsonElement("em_retvac")]
            public DateTime dtEm_retvac { get; set; }

            [BsonElement("em_fcreinf")]
            public DateTime dtEm_fcreinf { get; set; }

            [BsonElement("em_porbono")]
            public double fEm_porbono { get; set; }

            [BsonElement("em_moncre")]
            public double fEm_moncre { get; set; }

            [BsonElement("em_tipocta")]
            public string sEm_tipocta { get; set; }

            [BsonElement("em_tipoinf")]
            public string sEm_tipoinf { get; set; }

            [BsonElement("em_imssumf")]
            public double fEm_imssumf { get; set; }

            [BsonElement("em_deconom")]
            public double fEm_deconom { get; set; }

            [BsonElement("em_tabisr")]
            public string sEm_tabisr { get; set; }

            [BsonElement("em_ahdoble")]
            public int iEm_ahdoble { get; set; }

            [BsonElement("em_ahtrip")]
            public int iEm_ahtrip { get; set; }

            [BsonElement("em_adomica")]
            public int iEm_adomica { get; set; }

            [BsonElement("em_asabati")]
            public int iEm_asabati { get; set; }

            [BsonElement("em_aprod")]
            public int iEm_aprod { get; set; }

            [BsonElement("em_apuntal")]
            public int iEm_apuntal { get; set; }

            [BsonElement("em_ancalif")]
            public int iEm_ancalif { get; set; }

            [BsonElement("em_asisten")]
            public int iEm_asisten { get; set; }

            [BsonElement("em_turflot")]
            public int iEm_turflot { get; set; }

            [BsonElement("em_diasvaca")]
            public double fEm_diasvaca { get; set; }

            [BsonElement("em_porprima")]
            public double fEm_porprima { get; set; }
            #endregion

            #region nhumanos
            [BsonElement("rh_telefo")]
            public string sRh_telefo { get; set; }

            [BsonElement("rh_escolar")]
            public string sRh_escolar { get; set; }

            [BsonElement("rh_nciudad")]
            public string sRh_nciudad { get; set; }

            [BsonElement("rh_nestado")]
            public string sRh_nestado { get; set; }

            [BsonElement("rh_dcalle")]
            public string sRh_dcalle { get; set; }

            [BsonElement("rh_dcolon")]
            public string sRh_dcolon { get; set; }

            [BsonElement("rh_dciudad")]
            public string sRh_dciudad { get; set; }

            [BsonElement("rh_destado")]
            public string sRh_destado { get; set; }

            [BsonElement("rh_dmunici")]
            public string sRh_dmunici { get; set; }

            [BsonElement("rh_dcp")]
            public string sRh_dcp { get; set; }

            [BsonElement("rh_npadre")]
            public string sRh_npadre { get; set; }

            [BsonElement("rh_nmadre")]
            public string sRh_nmadre { get; set; }

            [BsonElement("rh_fpadre")]
            public int iRh_fpadre { get; set; }

            [BsonElement("rh_fmadre")]
            public int iRh_fmadre { get; set; }

            [BsonElement("rh_nacion")]
            public string sRh_nacion { get; set; }

            [BsonElement("rh_gmmsgv")]
            public string sRh_gmmsgv  { get; set; }

            [BsonElement("rh_gmm")]
            public string sRh_gmm { get; set; }

            [BsonElement("rh_gmmaseg")]
            public string sRh_gmmaseg { get; set; }

            [BsonElement("rh_gmmfnac")]
            public DateTime dtRh_gmmfnac { get; set; }

            [BsonElement("rh_gmmsexo")]
            public string sRh_gmmsexo { get; set; }

            [BsonElement("rh_gmmpare")]
            public string sRh_gmmpare { get; set; }

            [BsonElement("rh_gmmpor")]
            public double fRh_gmmpor { get; set; }

            [BsonElement("rh_noavis1")]
            public string sRh_noavis1 { get; set; }

            [BsonElement("rh_teavis1")]
            public string sRh_teavis1 { get; set; }

            [BsonElement("rh_paavis1")]
            public string sRh_paavis1 { get; set; }

            [BsonElement("rh_noavis2")]
            public string sRh_noavis2 { get; set; }

            [BsonElement("rh_teavis2")]
            public string sRh_teavis2 { get; set; }

            [BsonElement("rh_paavis2")]
            public string sRh_paavis2 { get; set; }

            [BsonElement("rh_dpenvac")]
            public double fRh_dpenvac { get; set; }

            [BsonElement("rh_picture")]
            public string sRh_picture { get; set; }

            [BsonElement("rh_gmmpcve")]
            public string sRh_gmmpcve { get; set; }

            [BsonElement("rh_area")]
            public double fRh_area { get; set; }

            [BsonElement("rh_oficio")]
            public string sRh_oficio { get; set; }

            [BsonElement("rh_gmmesta")]
            public double fRh_gmmesta { get; set; }

            [BsonElement("rh_gmmpeso")]
            public double fRh_gmmpeso { get; set; }

            [BsonElement("rh_plangmm")]
            public int iRh_plangmm { get; set; }

            [BsonElement("rh_plansv")]
            public int iRh_plansv { get; set; }

            [BsonElement("rh_gmmsuma")]
            public double fRh_gmmsuma { get; set; }

            [BsonElement("rh_aplansv")]
            public int iRh_aplansv { get; set; }

            [BsonElement("rh_psvsuma")]
            public double fRh_psvsuma { get; set; }

            [BsonElement("rh_ubicado")]
            public string sRh_ubicado { get; set; }

            [BsonElement("rh_estatu")]
            public double fRh_estatu { get; set; }

            [BsonElement("rh_peso")]
            public double fRh_peso { get; set; }

            [BsonElement("rh_tallac")]
            public double fRh_tallac { get; set; }

            [BsonElement("rh_tallap")]
            public double fRh_tallap { get; set; }

            [BsonElement("rh_calzado")]
            public double fRh_calzado { get; set; }

            [BsonElement("rh_coloroj")]
            public string sRh_coloroj { get; set; }

            [BsonElement("rh_colorca")]
            public string sRh_colorca { get; set; }

            [BsonElement("rh_piel")]
            public string sRh_piel { get; set; }

            [BsonElement("rh_separt")]
            public string sRh_separt { get; set; }

            [BsonElement("rh_svplane")]
            public double fRh_svplane { get; set; }

            [BsonElement("rh_deporte")]
            public string sRh_deporte { get; set; }

            [BsonElement("rh_hobbies")]
            public string sRh_hobbies { get; set; }

            [BsonElement("rh_lvacaci")]
            public string sRh_lvacaci { get; set; }

            [BsonElement("rh_numsmm")]
            public string sRh_numsmm { get; set; }

            [BsonElement("rh_numvisa")]
            public string sRh_numvisa { get; set; }

            [BsonElement("rh_numlic")]
            public string sRh_numlic { get; set; }

            [BsonElement("rh_benpens")]
            public string sRh_benpens { get; set; }

            [BsonElement("rh_piso")]
            public string sRh_piso { get; set; }

            [BsonElement("rh_teltra")]
            public string sRh_teltra { get; set; }

            [BsonElement("rh_exttrab")]
            public string sRh_exttrab { get; set; }

            [BsonElement("rh_extcerc")]
            public string sRh_extcerc { get; set; }

            [BsonElement("rh_secre")]
            public string sRh_secre { get; set; }

            [BsonElement("rh_telsec")]
            public string sRh_telsec { get; set; }

            [BsonElement("rh_extsec")]
            public string sRh_extsec { get; set; }

            [BsonElement("rh_mail")]
            public string sRh_mail { get; set; }

            [BsonElement("rh_benpen2")]
            public string sRh_benpen2 { get; set; }

            [BsonElement("rh_fecapti")]
            public DateTime dtRh_fecapti { get; set; }

            [BsonElement("rh_fecflot")]
            public DateTime dtRh_fecflot { get; set; }

            [BsonElement("rh_fecind")]
            public DateTime dtRh_fecind { get; set; }

            [BsonElement("rh_svcosto")]
            public double fRh_svcosto { get; set; }

            [BsonElement("rh_gmcosto")]
            public double fRh_gmcosto { get; set; }

            [BsonElement("rh_soceco")]
            public DateTime dtRh_soceco { get; set; }

            [BsonElement("rh_segpub")]
            public DateTime dtRh_segpub { get; set; }

            [BsonElement("rh_antidop")]
            public DateTime dtRh_antidop { get; set; }

            [BsonElement("rh_cartila")]
            public string sRh_cartila { get; set; }
            #endregion

            #region nestudio
            [BsonElement("em_ulgrado")]
            public string sEm_ulgrado { get; set; }

            [BsonElement("em_uinsti")]
            public string sEm_uinsti { get; set; }

            [BsonElement("em_genera")]
            public string sEm_genera { get; set; }


            [BsonElement("em_utitulo")]
            public string sEm_utitulo { get; set; }

            [BsonElement("em_maestri")]
            public string sEm_maestri { get; set; }

            [BsonElement("em_minsti")]
            public string sEm_minsti { get; set; }

            [BsonElement("em_docto")]
            public string sEm_docto { get; set; }

            [BsonElement("em_dinsti")]
            public string sEm_dinsti { get; set; }

            [BsonElement("em_especia")]
            public string sEm_especia { get; set; }

            [BsonElement("em_einsti")]
            public string sEm_einsti { get; set; }

            [BsonElement("em_cedula")]
            public string sEm_cedula { get; set; }

            [BsonElement("em_titsn")]
            public string sEm_titsn { get; set; }

            [BsonElement("em_diplomado")]
            public string sEm_diplomado { get; set; }

            [BsonElement("em_observ")]
            public string sEm_observ { get; set; }
            #endregion

            //Tabla nbenefi
            [BsonElement("em_beneficiarios")]
            public mdl_Beneficiario[] aBeneficiarios { get; set; }


            //Tabla nbajas 
            [BsonElement("em_bajas")]
            public mdl_Baja[] aBajas { get; set; }


            //Tabla nemprein
            [BsonElement("em_reingresos")]
            public mdl_Reingreso[] aReingreso { get; set; }


            //Tabla nemphist
            [BsonElement("em_historico")]
            public mdl_Historico[] aHistorico { get; set; }


            //Tabla nsalario
            [BsonElement("em_incrementos")]
            public mdl_Salario[] aIncremento { get; set; }


            //Tabla hvacthis
            [BsonElement("em_vacaciones")]
            public mdl_Vacaciones[] aVacaciones { get; set; }

            //tabla nemplepd
            [BsonElement("em_conceptos")]
            public mdl_Conceptos[] aConceptos { get; set; }


            [BsonElement("sr_recno")]
            public long iSr_recno { get; set; }

            [BsonElement("sr_deleted")]
            public string sSr_Deleted { get; set; }
        }
        /// <summary>
        /// Clase que contiene los elementos del arreglo beneficiarios, correspondiente a la tabla nbenefi (Clase NEmplea)
        /// </summary>
        public class mdl_Beneficiario
        {
            [BsonElement("be_tipo")]
            public string sBe_tipo { get; set; }

            [BsonElement("be_clavep")]
            public string sBe_clavep { get; set; }

            [BsonElement("be_nombre")]
            public string sBe_nombre { get; set; }

            [BsonElement("be_parient")]
            public string sBe_parient { get; set; }

            [BsonElement("be_porcent")]
            public double fBe_porcent { get; set; }

            [BsonElement("be_fecnac")]
            public DateTime dtBe_fecnac { get; set; }

            [BsonElement("be_sexo")]
            public string sBe_sexo { get; set; }

            [BsonElement("be_telefo")]
            public string sBe_telefo { get; set; }

            [BsonElement("be_peso")]
            public double fBe_peso { get; set; }

            [BsonElement("be_estatur")]
            public double fBe_estatur { get; set; }

            [BsonElement("be_estudio")]
            public string sBe_estudio { get; set; }

            [BsonElement("be_deporte")]
            public string sBe_deporte { get; set; }

            [BsonElement("be_hobbies")]
            public string sBe_hobbies { get; set; }
        }

        /// <summary>
        /// Clase que contiene los elementos del arreglo de bajas, correspondiente a la tabla nbajas (Clase NEmplea)
        /// </summary>
        public class mdl_Baja
        {

            [BsonElement("em_bfechai")]
            public DateTime dtEm_bfechai { get; set; }

            [BsonElement("em_bfechab")]
            public DateTime dtEm_bfechab { get; set; }

            [BsonElement("em_bcausa")]
            public string sEm_bcausa { get; set; }

            [BsonElement("em_comenta")]
            public string sEm_comenta { get; set; }

            [BsonElement("em_salariob")]
            public double fEm_salariob { get; set; }

            [BsonElement("em_salinte")]
            public double fEm_salinte { get; set; }

            [BsonElement("em_sdivar")]
            public double fEm_sdivar { get; set; }

            [BsonElement("em_bajaim")]
            public DateTime dtEm_bajaim { get; set; }

            [BsonElement("em_bstatimp")]
            public int iEm_bstatimp { get; set; }

            [BsonElement("em_baimss")]
            public string sEm_baimss { get; set; }

            [BsonElement("em_norecon")]
            public int iEm_norecon { get; set; }

            [BsonElement("em_bdepto")]
            public int iEm_bdepto { get; set; }

            [BsonElement("em_bpuesto")]
            public int iEm_bpuesto { get; set; }

        }

        /// <summary>
        /// Clase que contiene los elementos del arreglo reingresos, correspondiente a la tabla nbreingre (Clase NEmplea)
        /// </summary>
        public class mdl_Reingreso
        {
            [BsonElement("em_rfecha")]
            public DateTime dtEm_rfecha { get; set; }

            [BsonElement("em_sdifijo")]
            public double fEm_sdifijo { get; set; }

            [BsonElement("em_rsdivar")]
            public double fEm_rsdivar { get; set; }

            [BsonElement("em_rcausa")]
            public string sEm_rcausa { get; set; }

            [BsonElement("em_fechaba")]
            public DateTime dtEm_fechaba { get; set; }

            [BsonElement("em_rsalario")]
            public double fEm_rsalario { get; set; }

            [BsonElement("em_rfecnac")]
            public DateTime dtEm_rfecnac { get; set; }
        }

        /// <summary>
        /// Clase que contiene los elementos del arreglo historico, correspondiente a la tabla nemphist (Clase NEmplea)
        /// </summary>
        public class mdl_Historico
        {
            [BsonElement("em_campo")]
            public string sEm_campo { get; set; }

            [BsonElement("em_cambio")]
            public DateTime dtEm_cambio { get; set; }

            [BsonElement("em_tipomov")]
            public string sEm_tipomov { get; set; }

            [BsonElement("em_quecam")]
            public string sEm_quecam { get; set; }

            [BsonElement("em_valant")]
            public string sEm_valant { get; set; }

            [BsonElement("em_valnew")]
            public string sEm_valnew { get; set; }
        }

        /// <summary>
        /// Clase que contiene los elementos del arreglo incrementos, correspondiente a la tabla nsalarios (Clase NEmplea)
        /// </summary>
        public class mdl_Salario
        {
            [BsonElement("em_salant")]
            public double fEm_salant { get; set; }

            [BsonElement("em_fecant")]
            public DateTime dtEm_fecant { get; set; }

            [BsonElement("em_salnew")]
            public double fEm_salnew { get; set; }

            [BsonElement("em_feccam")]
            public DateTime dtEm_feccam { get; set; }

            [BsonElement("em_intant")]
            public double fEm_intant { get; set; }

            [BsonElement("em_intnew")]
            public double fEm_intnew { get; set; }

            [BsonElement("em_sdivant")]
            public double fEm_sdivant { get; set; }

            [BsonElement("em_nivela")]
            public int iEm_nivela { get; set; }

            [BsonElement("em_puestoa")]
            public int iEm_puestoa { get; set; }

            [BsonElement("em_topes")]
            public int iEm_topes { get; set; }

            [BsonElement("em_istatimp")]
            public int iEm_istatimp { get; set; }

            [BsonElement("em_sdianiv")]
            public int iEm_sdianiv { get; set; }
        }

        /// <summary>
        /// Clase que contiene los elementos del arreglo vacaciones, correspondiente a la tabla hvacthis (Clase NEmplea)
        /// </summary>
        public class mdl_Vacaciones
        {

            [BsonElement("em_fechaini")]
            public DateTime dtEm_fechaini { get; set; }

            [BsonElement("em_vsuper")]
            public int iEm_vsuper { get; set; }

            [BsonElement("em_vdepto")]
            public int iEm_vdepto { get; set; }

            [BsonElement("em_vsucursa")]
            public int iEm_vsucursa { get; set; }

            [BsonElement("em_vsaldia")]
            public double fEm_vsaldia { get; set; }

            [BsonElement("vh_anio")]
            public int iVh_anio { get; set; }

            [BsonElement("vh_ciclo")]
            public string sVh_ciclo { get; set; }

            [BsonElement("vh_diasder")]
            public int iVh_diasder { get; set; }

            [BsonElement("vh_disfrut")]
            public int iVh_disfrut { get; set; }

            [BsonElement("vh_porprim")]
            public double fVh_porprim { get; set; }

            [BsonElement("vh_diaspen")]
            public int iVh_diaspen { get; set; }
        }

        /// <summary>
        /// Clase que contiene los elementos del arreglo conceptos, correspondiente a la tabla nemplepd (Clase NEmplea)
        /// </summary>
        public class mdl_Conceptos
        {
            [BsonElement("pd_periodo")]
            public int iPd_periodo { get; set; }

            [BsonElement("pd_tiponom")]
            public string sPd_tiponom { get; set; }

            [BsonElement("pd_tipoemp")]
            public string sPd_tipoemp { get; set; }

            [BsonElement("pd_tipocon")]
            public string sPd_tipocon { get; set; }

            [BsonElement("pd_numero")]
            public int iPd_numero { get; set; }

            [BsonElement("pd_monto")]
            public double fPd_monto { get; set; }

            [BsonElement("pd_tiempo")]
            public string sPd_tiempo { get; set; }

            [BsonElement("pd_modelo")]
            public string sPd_modelo { get; set; }

            [BsonElement("pd_activi")]
            public string sPd_activi { get; set; }
        }
    }
}
