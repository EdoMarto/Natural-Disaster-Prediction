using OfficeOpenXml;
using System.Data.SqlClient;

namespace Thesis
{
    public class File16
    {
        public static void run()
        {
            Console.WriteLine("Sto avviando File16...");
            int n_inseriti = manageExcel("Indicatori_Intero_territorio_nazionale.xlsx");
            Console.WriteLine("Esecuzione File16 terminata. Sono stati inseriti " + n_inseriti + " nuovi record");
        }

        static void insertDB(object tmp1, object tmp2, object tmp3, object tmp4, object tmp5, object tmp6, object tmp7, object tmp8, object tmp9, object tmp10, object tmp11, object tmp12, object tmp13, object tmp14, object tmp15, object tmp16, object tmp17, object tmp18, object tmp19, object tmp20, object tmp21, object tmp22, object tmp23, object tmp24, object tmp25, object tmp26, object tmp27, object tmp28, object tmp29, object tmp30, object tmp31, object tmp32, object tmp33, object tmp34, object tmp35, object tmp36, object tmp37, object tmp38, object tmp39, object tmp40, object tmp41, object tmp42, object tmp43, object tmp44, object tmp45, object tmp46, object tmp47, object tmp48, object tmp49, object tmp50, object tmp51, object tmp52, object tmp53, object tmp54, object tmp55, object tmp56, object tmp57, object tmp58, object tmp59, object tmp60, object tmp61, object tmp62, object tmp63, object tmp64, object tmp65, object tmp66, object tmp67, object tmp68, object tmp69, object tmp70, object tmp71, object tmp72, object tmp73, object tmp74, object tmp75, object tmp76, object tmp77, object tmp78, object tmp79, object tmp80, object tmp81, object tmp82, object tmp83, object tmp84, object tmp85, object tmp86, object tmp87, object tmp88, object tmp89, object tmp90, object tmp91, object tmp92, object tmp93, object tmp94, object tmp95, object tmp96, object tmp97, object tmp98, object tmp99, object tmp100, object tmp101, object tmp102, object tmp103, object tmp104, object tmp105, object tmp106, object tmp107, object tmp108, object tmp109, object tmp110, object tmp111, object tmp112, object tmp113, object tmp114, object tmp115, object tmp116, object tmp117, object tmp118, object tmp119, object tmp120, object tmp121, object tmp122, object tmp123, object tmp124, object tmp125, object tmp126, object tmp127, object tmp128, object tmp129, object tmp130, object tmp131, object tmp132, object tmp133, object tmp134, object tmp135, object tmp136, object tmp137, object tmp138, object tmp139, object tmp140, object tmp141, object tmp142, object tmp143, object tmp144, object tmp145, object tmp146, object tmp147, object tmp148, object tmp149, object tmp150, object tmp151, object tmp152, object tmp153, object tmp154, object tmp155, object tmp156, object tmp157, object tmp158, object tmp159, object tmp160, object tmp161, object tmp162, object tmp163, object tmp164, object tmp165, object tmp166, object tmp167, object tmp168, object tmp169, object tmp170, object tmp171, object tmp172, object tmp173, object tmp174, object tmp175, object tmp176, object tmp177, object tmp178, object tmp179, object tmp180, object tmp181, object tmp182, object tmp183, object tmp184, object tmp185, object tmp186, object tmp187, object tmp188, object tmp189, object tmp190, object tmp191, object tmp192, object tmp193, object tmp194, object tmp195, object tmp196, object tmp197, object tmp198, object tmp199, object tmp200, object tmp201, object tmp202, object tmp203, object tmp204, object tmp205, object tmp206, object tmp207, object tmp208, object tmp209, object tmp210, object tmp211, object tmp212, object tmp213, object tmp214, object tmp215, object tmp216, object tmp217, object tmp218, object tmp219, object tmp220, object tmp221, object tmp222, object tmp223, object tmp224, object tmp225, object tmp226, object tmp227, object tmp228, object tmp229, object tmp230, object tmp231, object tmp232, object tmp233, object tmp234, object tmp235, object tmp236, object tmp237, object tmp238, object tmp239, object tmp240, object tmp241, object tmp242, object tmp243, object tmp244, object tmp245, object tmp246, object tmp247, object tmp248, object tmp249, object tmp250, object tmp251, object tmp252, object tmp253, object tmp254, object tmp255, object tmp256, object tmp257, object tmp258, object tmp259, object tmp260, object tmp261, object tmp262, object tmp263, object tmp264, object tmp265, object tmp266, object tmp267, object tmp268, object tmp269, object tmp270, object tmp271, object tmp272, object tmp273, object tmp274, object tmp275, object tmp276, object tmp277, object tmp278, object tmp279)
        {
            string connectionString = Settings.connection_string;
            SqlConnection connection = new SqlConnection(@connectionString);
            string query = "INSERT INTO tbl_13029_Stage_Istat (ALT_LOC_AB, DataInserimento, DataUltimaModifica, Deleted, A2, A2_P, A2_R, A3, A3_P, A3_R, A7, A7_P, A7_R, COMUNE_MONTANO, COMUNE_MONTANO_TXT, DENSPOP, DENSPOP_P, DENSPOP_R, PERC_ECP3_8_12, PERC_ECP3_13_15, PERC_ECP3_16, PERC_ECP3_8_12_P, PERC_ECP3_13_15_P, PERC_ECP3_16_P, PERC_ECP3_8_12_R, PERC_ECP3_13_15_R, PERC_ECP3_16_R, PERC_EC8_12, PERC_EC16, PERC_EC13_15, PERC_EC8_12_P, PERC_EC13_15_P, PERC_EC16_P, PERC_EC8_12_R, PERC_EC13_15_R, PERC_EC16_R, PERC_EMC_O, PERC_EMC_O_P, PERC_EMC_O_R, PERC_EMC_B, PERC_EMC_B_P, PERC_EMC_B_R, PERC_EMC_M, PERC_EMC_M_P, PERC_EMC_M_R, PERC_EMC_P, PERC_EMC_P_P, PERC_EMC_P_R, PERC_EME8, PERC_EME8_P, PERC_EME8_R, PERC_EME9, PERC_EME9_P, PERC_EME9_R, PERC_EME10, PERC_EME10_P, PERC_EME10_R, PERC_EME11, PERC_EME11_P, PERC_EME11_R, PERC_EME12, PERC_EME12_P, PERC_EME12_R, PERC_EME13, PERC_EME13_P, PERC_EME13_R, PERC_EME14, PERC_EME14_P, PERC_EME14_R, PERC_EME15, PERC_EME15_P, PERC_EME15_R, PERC_EME16, PERC_EME16_P, PERC_EME16_R, PERC_EMP1, PERC_EMP1_P, PERC_EMP1_R, PERC_EMP2, PERC_EMP2_P, PERC_EMP2_R, PERC_EMP3, PERC_EMP3_P, PERC_EMP3_R, PERC_EP3C, PERC_EP3C_P, PERC_EP3C_R, PERC_EP3C1, PERC_EP3C1_P, PERC_EP3C1_R, PERC_EP3C2, PERC_EP3C2_P, PERC_EP3C2_R, PERC_EP3C3, PERC_EP3C3_P, PERC_EP3C3_R, PERC_EP3C4, PERC_EP3C4_P, PERC_EP3C4_R, PERC_EP3E1, PERC_EP3E1_P, PERC_EP3E1_R, PERC_EP3E2, PERC_EP3E2_P, PERC_EP3E2_R, PERC_EP3E3, PERC_EP3E3_P, PERC_EP3E3_R, PERC_EP3E4, PERC_EP3E4_P, PERC_EP3E4_R, PERC_EP3E5, PERC_EP3E5_P, PERC_EP3E5_R, PERC_EP3E6, PERC_EP3E6_P, PERC_EP3E6_R, PERC_EP3E7, PERC_EP3E7_P, PERC_EP3E7_R, PERC_EP3E8, PERC_EP3E8_P, PERC_EP3E8_R, PERC_EP3E9, PERC_EP3E9_P, PERC_EP3E9_R, PERC_EP3M, PERC_EP3M_P, PERC_EP3M_R, PERC_ERA1, PERC_ERA1_P, PERC_ERA1_R, PERC_ERC1, PERC_ERC1_P, PERC_ERC1_R, ERECEM, ERECQ, ERECQ_TXT, EREMP, ERESPP, PERC_ERE8, PERC_ERE8_P, PERC_ERE8_R, PERC_ERE9, PERC_ERE9_P, PERC_ERE9_R, PERC_ERE10, PERC_ERE10_P, PERC_ERE10_R, PERC_ERE11, PERC_ERE11_P, PERC_ERE11_R, PERC_ERE12, PERC_ERE12_P, PERC_ERE12_R, PERC_ERE13, PERC_ERE13_P, PERC_ERE13_R, PERC_ERE14, PERC_ERE14_P, PERC_ERE14_R, PERC_ERE15, PERC_ERE15_P, PERC_ERE15_R, PERC_ERE16, PERC_ERE16_P, PERC_ERE16_R, PERC_ERM1, PERC_ERM1_P, PERC_ERM1_R, PERC_ETA_Q1, PERC_ETA_Q1_P, PERC_ETA_Q1_R, PERC_ETA_Q2, PERC_ETA_Q2_P, PERC_ETA_Q2_R, PERC_ETA_Q3, PERC_ETA_Q3_P, PERC_ETA_Q3_R, PERC_ETA_Q4, PERC_ETA_Q4_P, PERC_ETA_Q4_R, E3, E3_P, E3_R, E5, E5_P, E5_R, E6, E6_P, E6_R, E7, E7_P, E7_R, E8, E8_P, E8_R, E9, E9_P, E9_R, E10, E10_P, E10_R, E11, E11_P, E11_R, E12, E12_P, E12_R, E13, E13_P, E13_R, E14, E14_P, E14_R, E15, E15_P, E15_R, E16, E16_P, E16_R, E17, E17_P, E17_R, E18, E18_P, E18_R, E19, E19_P, E19_R, E20, E20_P, E20_R, E28, E28_P, E28_R, E29, E29_P, E29_R, E30, E30_P, E30_R, E31, E31_P, E31_R, FAM_2011, FAM_2011_P, FAM_2011_R, FAM_2018, FAM_2018_P, FAM_2018_R, IDEM, IDEM_P, IDEM_R, PERC_IND_DIP_STR, PERC_IND_DIP_STR_P, PERC_IND_DIP_STR_R, PERC_IVSM, LIT, LIT_TXT, PERC_POP_ANZ, PERC_POP_ANZ_P, PERC_POP_ANZ_R, POP_2011, POP_2011_P, POP_2011_R, POP_2018, POP_2018_P, POP_2018_R, PRO_COM_110, SUP, SUP_P, SUP_R, SUP_URB, SUP_URB_P, SUP_URB_R, PERC_VAR_PERC, PERC_VAR_PERC_P, PERC_VAR_PERC_R, VECCH, VECCH_P, VECCH_R, DataImportazione, ParentWeb_Id) " +
                "VALUES(@ALT_LOC_AB, @DataInserimento, @DataUltimaModifica, @Deleted, @A2, @A2_P, @A2_R, @A3, @A3_P, @A3_R, @A7, @A7_P, @A7_R, @COMUNE_MONTANO, @COMUNE_MONTANO_TXT, @DENSPOP, @DENSPOP_P, @DENSPOP_R, @PERC_ECP3_8_12, @PERC_ECP3_13_15, @PERC_ECP3_16, @PERC_ECP3_8_12_P, @PERC_ECP3_13_15_P, @PERC_ECP3_16_P, @PERC_ECP3_8_12_R, @PERC_ECP3_13_15_R, @PERC_ECP3_16_R, @PERC_EC8_12, @PERC_EC16, @PERC_EC13_15, @PERC_EC8_12_P, @PERC_EC13_15_P, @PERC_EC16_P, @PERC_EC8_12_R, @PERC_EC13_15_R, @PERC_EC16_R, @PERC_EMC_O, @PERC_EMC_O_P, @PERC_EMC_O_R, @PERC_EMC_B, @PERC_EMC_B_P, @PERC_EMC_B_R, @PERC_EMC_M, @PERC_EMC_M_P, @PERC_EMC_M_R, @PERC_EMC_P, @PERC_EMC_P_P, @PERC_EMC_P_R, @PERC_EME8, @PERC_EME8_P, @PERC_EME8_R, @PERC_EME9, @PERC_EME9_P, @PERC_EME9_R, @PERC_EME10, @PERC_EME10_P, @PERC_EME10_R, @PERC_EME11, @PERC_EME11_P, @PERC_EME11_R, @PERC_EME12, @PERC_EME12_P, @PERC_EME12_R, @PERC_EME13, @PERC_EME13_P, @PERC_EME13_R, @PERC_EME14, @PERC_EME14_P, @PERC_EME14_R, @PERC_EME15, @PERC_EME15_P, @PERC_EME15_R, @PERC_EME16, @PERC_EME16_P, @PERC_EME16_R, @PERC_EMP1, @PERC_EMP1_P, @PERC_EMP1_R, @PERC_EMP2, @PERC_EMP2_P, @PERC_EMP2_R, @PERC_EMP3, @PERC_EMP3_P, @PERC_EMP3_R, @PERC_EP3C, @PERC_EP3C_P, @PERC_EP3C_R, @PERC_EP3C1, @PERC_EP3C1_P, @PERC_EP3C1_R, @PERC_EP3C2, @PERC_EP3C2_P, @PERC_EP3C2_R, @PERC_EP3C3, @PERC_EP3C3_P, @PERC_EP3C3_R, @PERC_EP3C4, @PERC_EP3C4_P, @PERC_EP3C4_R, @PERC_EP3E1, @PERC_EP3E1_P, @PERC_EP3E1_R, @PERC_EP3E2, @PERC_EP3E2_P, @PERC_EP3E2_R, @PERC_EP3E3, @PERC_EP3E3_P, @PERC_EP3E3_R, @PERC_EP3E4, @PERC_EP3E4_P, @PERC_EP3E4_R, @PERC_EP3E5, @PERC_EP3E5_P, @PERC_EP3E5_R, @PERC_EP3E6, @PERC_EP3E6_P, @PERC_EP3E6_R, @PERC_EP3E7, @PERC_EP3E7_P, @PERC_EP3E7_R, @PERC_EP3E8, @PERC_EP3E8_P, @PERC_EP3E8_R, @PERC_EP3E9, @PERC_EP3E9_P, @PERC_EP3E9_R, @PERC_EP3M, @PERC_EP3M_P, @PERC_EP3M_R, @PERC_ERA1, @PERC_ERA1_P, @PERC_ERA1_R, @PERC_ERC1, @PERC_ERC1_P, @PERC_ERC1_R, @ERECEM, @ERECQ, @ERECQ_TXT, @EREMP, @ERESPP, @PERC_ERE8, @PERC_ERE8_P, @PERC_ERE8_R, @PERC_ERE9, @PERC_ERE9_P, @PERC_ERE9_R, @PERC_ERE10, @PERC_ERE10_P, @PERC_ERE10_R, @PERC_ERE11, @PERC_ERE11_P, @PERC_ERE11_R, @PERC_ERE12, @PERC_ERE12_P, @PERC_ERE12_R, @PERC_ERE13, @PERC_ERE13_P, @PERC_ERE13_R, @PERC_ERE14, @PERC_ERE14_P, @PERC_ERE14_R, @PERC_ERE15, @PERC_ERE15_P, @PERC_ERE15_R, @PERC_ERE16, @PERC_ERE16_P, @PERC_ERE16_R, @PERC_ERM1, @PERC_ERM1_P, @PERC_ERM1_R, @PERC_ETA_Q1, @PERC_ETA_Q1_P, @PERC_ETA_Q1_R, @PERC_ETA_Q2, @PERC_ETA_Q2_P, @PERC_ETA_Q2_R, @PERC_ETA_Q3, @PERC_ETA_Q3_P, @PERC_ETA_Q3_R, @PERC_ETA_Q4, @PERC_ETA_Q4_P, @PERC_ETA_Q4_R, @E3, @E3_P, @E3_R, @E5, @E5_P, @E5_R, @E6, @E6_P, @E6_R, @E7, @E7_P, @E7_R, @E8, @E8_P, @E8_R, @E9, @E9_P, @E9_R, @E10, @E10_P, @E10_R, @E11, @E11_P, @E11_R, @E12, @E12_P, @E12_R, @E13, @E13_P, @E13_R, @E14, @E14_P, @E14_R, @E15, @E15_P, @E15_R, @E16, @E16_P, @E16_R, @E17, @E17_P, @E17_R, @E18, @E18_P, @E18_R, @E19, @E19_P, @E19_R, @E20, @E20_P, @E20_R, @E28, @E28_P, @E28_R, @E29, @E29_P, @E29_R, @E30, @E30_P, @E30_R, @E31, @E31_P, @E31_R, @FAM_2011, @FAM_2011_P, @FAM_2011_R, @FAM_2018, @FAM_2018_P, @FAM_2018_R, @IDEM, @IDEM_P, @IDEM_R, @PERC_IND_DIP_STR, @PERC_IND_DIP_STR_P, @PERC_IND_DIP_STR_R, @PERC_IVSM, @LIT, @LIT_TXT, @PERC_POP_ANZ, @PERC_POP_ANZ_P, @PERC_POP_ANZ_R, @POP_2011, @POP_2011_P, @POP_2011_R, @POP_2018, @POP_2018_P, @POP_2018_R, @PRO_COM_110, @SUP, @SUP_P, @SUP_R, @SUP_URB, @SUP_URB_P, @SUP_URB_R, @PERC_VAR_PERC, @PERC_VAR_PERC_P, @PERC_VAR_PERC_R, @VECCH, @VECCH_P, @VECCH_R, @DataImportazione, @ParentWeb_Id)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@ALT_LOC_AB", tmp1 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DataInserimento", DateTime.Now);
            command.Parameters.AddWithValue("@DataUltimaModifica", DateTime.Now);
            command.Parameters.AddWithValue("@Deleted", false);
            command.Parameters.AddWithValue("@A2", tmp2 ?? DBNull.Value);
            command.Parameters.AddWithValue("@A2_P", tmp3 ?? DBNull.Value);
            command.Parameters.AddWithValue("@A2_R", tmp4 ?? DBNull.Value);
            command.Parameters.AddWithValue("@A3", tmp5 ?? DBNull.Value);
            command.Parameters.AddWithValue("@A3_P", tmp6 ?? DBNull.Value);
            command.Parameters.AddWithValue("@A3_R", tmp7 ?? DBNull.Value);
            command.Parameters.AddWithValue("@A7", tmp8 ?? DBNull.Value);
            command.Parameters.AddWithValue("@A7_P", tmp9 ?? DBNull.Value);
            command.Parameters.AddWithValue("@A7_R", tmp10 ?? DBNull.Value);
            command.Parameters.AddWithValue("@COMUNE_MONTANO", tmp11 ?? DBNull.Value);
            command.Parameters.AddWithValue("@COMUNE_MONTANO_TXT", tmp12 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DENSPOP", tmp13 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DENSPOP_P", tmp14 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DENSPOP_R", tmp15 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_8_12", tmp16 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_13_15", tmp17 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_16", tmp18 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_8_12_P", tmp19 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_13_15_P", tmp20 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_16_P", tmp21 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_8_12_R", tmp22 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_13_15_R", tmp23 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ECP3_16_R", tmp24 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC8_12", tmp25 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC16", tmp26 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC13_15", tmp27 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC8_12_P", tmp28 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC13_15_P", tmp29 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC16_P", tmp30 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC8_12_R", tmp31 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC13_15_R", tmp32 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EC16_R", tmp33 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_O", tmp34 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_O_P", tmp35 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_O_R", tmp36 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_B", tmp37 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_B_P", tmp38 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_B_R", tmp39 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_M", tmp40 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_M_P", tmp41 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_M_R", tmp42 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_P", tmp43 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_P_P", tmp44 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMC_P_R", tmp45 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME8", tmp46 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME8_P", tmp47 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME8_R", tmp48 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME9", tmp49 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME9_P", tmp50 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME9_R", tmp51 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME10", tmp52 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME10_P", tmp53 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME10_R", tmp54 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME11", tmp55 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME11_P", tmp56 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME11_R", tmp57 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME12", tmp58 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME12_P", tmp59 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME12_R", tmp60 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME13", tmp61 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME13_P", tmp62 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME13_R", tmp63 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME14", tmp64 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME14_P", tmp65 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME14_R", tmp66 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME15", tmp67 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME15_P", tmp68 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME15_R", tmp69 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME16", tmp70 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME16_P", tmp71 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EME16_R", tmp72 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP1", tmp73 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP1_P", tmp74 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP1_R", tmp75 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP2", tmp76 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP2_P", tmp77 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP2_R", tmp78 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP3", tmp79 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP3_P", tmp80 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EMP3_R", tmp81 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C", tmp82 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C_P", tmp83 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C_R", tmp84 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C1", tmp85 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C1_P", tmp86 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C1_R", tmp87 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C2", tmp88 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C2_P", tmp89 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C2_R", tmp90 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C3", tmp91 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C3_P", tmp92 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C3_R", tmp93 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C4", tmp94 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C4_P", tmp95 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3C4_R", tmp96 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E1", tmp97 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E1_P", tmp98 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E1_R", tmp99 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E2", tmp100 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E2_P", tmp101 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E2_R", tmp102 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E3", tmp103 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E3_P", tmp104 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E3_R", tmp105 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E4", tmp106 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E4_P", tmp107 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E4_R", tmp108 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E5", tmp109 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E5_P", tmp110 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E5_R", tmp111 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E6", tmp112 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E6_P", tmp113 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E6_R", tmp114 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E7", tmp115 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E7_P", tmp116 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E7_R", tmp117 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E8", tmp118 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E8_P", tmp119 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E8_R", tmp120 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E9", tmp121 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E9_P", tmp122 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3E9_R", tmp123 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3M", tmp124 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3M_P", tmp125 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_EP3M_R", tmp126 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERA1", tmp127 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERA1_P", tmp128 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERA1_R", tmp129 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERC1", tmp130 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERC1_P", tmp131 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERC1_R", tmp132 ?? DBNull.Value);
            command.Parameters.AddWithValue("@ERECEM", tmp133 ?? DBNull.Value);
            command.Parameters.AddWithValue("@ERECQ", tmp134 ?? DBNull.Value);
            command.Parameters.AddWithValue("@ERECQ_TXT", tmp135 ?? DBNull.Value);
            command.Parameters.AddWithValue("@EREMP", tmp136 ?? DBNull.Value);
            command.Parameters.AddWithValue("@ERESPP", tmp137 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE8", tmp138 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE8_P", tmp139 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE8_R", tmp140 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE9", tmp141 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE9_P", tmp142 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE9_R", tmp143 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE10", tmp144 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE10_P", tmp145 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE10_R", tmp146 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE11", tmp147 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE11_P", tmp148 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE11_R", tmp149 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE12", tmp150 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE12_P", tmp151 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE12_R", tmp152 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE13", tmp153 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE13_P", tmp154 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE13_R", tmp155 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE14", tmp156 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE14_P", tmp157 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE14_R", tmp158 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE15", tmp159 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE15_P", tmp160 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE15_R", tmp161 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE16", tmp162 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE16_P", tmp163 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERE16_R", tmp164 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERM1", tmp165 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERM1_P", tmp166 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ERM1_R", tmp167 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q1", tmp168 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q1_P", tmp169 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q1_R", tmp170 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q2", tmp171 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q2_P", tmp172 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q2_R", tmp173 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q3", tmp174 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q3_P", tmp175 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q3_R", tmp176 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q4", tmp177 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q4_P", tmp178 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_ETA_Q4_R", tmp179 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E3", tmp180 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E3_P", tmp181 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E3_R", tmp182 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E5", tmp183 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E5_P", tmp184 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E5_R", tmp185 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E6", tmp186 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E6_P", tmp187 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E6_R", tmp188 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E7", tmp189 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E7_P", tmp190 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E7_R", tmp191 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E8", tmp192 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E8_P", tmp193 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E8_R", tmp194 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E9", tmp195 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E9_P", tmp196 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E9_R", tmp197 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E10", tmp198 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E10_P", tmp199 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E10_R", tmp200 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E11", tmp201 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E11_P", tmp202 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E11_R", tmp203 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E12", tmp204 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E12_P", tmp205 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E12_R", tmp206 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E13", tmp207 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E13_P", tmp208 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E13_R", tmp209 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E14", tmp210 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E14_P", tmp211 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E14_R", tmp212 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E15", tmp213 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E15_P", tmp214 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E15_R", tmp215 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E16", tmp216 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E16_P", tmp217 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E16_R", tmp218 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E17", tmp219 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E17_P", tmp220 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E17_R", tmp221 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E18", tmp222 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E18_P", tmp223 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E18_R", tmp224 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E19", tmp225 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E19_P", tmp226 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E19_R", tmp227 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E20", tmp228 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E20_P", tmp229 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E20_R", tmp230 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E28", tmp231 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E28_P", tmp232 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E28_R", tmp233 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E29", tmp234 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E29_P", tmp235 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E29_R", tmp236 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E30", tmp237 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E30_P", tmp238 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E30_R", tmp239 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E31", tmp240 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E31_P", tmp241 ?? DBNull.Value);
            command.Parameters.AddWithValue("@E31_R", tmp242 ?? DBNull.Value);
            command.Parameters.AddWithValue("@FAM_2011", tmp243 ?? DBNull.Value);
            command.Parameters.AddWithValue("@FAM_2011_P", tmp244 ?? DBNull.Value);
            command.Parameters.AddWithValue("@FAM_2011_R", tmp245 ?? DBNull.Value);
            command.Parameters.AddWithValue("@FAM_2018", tmp246 ?? DBNull.Value);
            command.Parameters.AddWithValue("@FAM_2018_P", tmp247 ?? DBNull.Value);
            command.Parameters.AddWithValue("@FAM_2018_R", tmp248 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDEM", tmp249 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDEM_P", tmp250 ?? DBNull.Value);
            command.Parameters.AddWithValue("@IDEM_R", tmp251 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_IND_DIP_STR", tmp52 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_IND_DIP_STR_P", tmp253 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_IND_DIP_STR_R", tmp254 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_IVSM", tmp255 ?? DBNull.Value);
            command.Parameters.AddWithValue("@LIT", tmp256 ?? DBNull.Value);
            command.Parameters.AddWithValue("@LIT_TXT", tmp257 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_POP_ANZ", tmp258 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_POP_ANZ_P", tmp259 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_POP_ANZ_R", tmp260 ?? DBNull.Value);
            command.Parameters.AddWithValue("@POP_2011", tmp261 ?? DBNull.Value);
            command.Parameters.AddWithValue("@POP_2011_P", tmp262 ?? DBNull.Value);
            command.Parameters.AddWithValue("@POP_2011_R", tmp263 ?? DBNull.Value);
            command.Parameters.AddWithValue("@POP_2018", tmp264 ?? DBNull.Value);
            command.Parameters.AddWithValue("@POP_2018_P", tmp265 ?? DBNull.Value);
            command.Parameters.AddWithValue("@POP_2018_R", tmp266 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PRO_COM_110", tmp267 ?? DBNull.Value);
            command.Parameters.AddWithValue("@SUP", tmp268 ?? DBNull.Value);
            command.Parameters.AddWithValue("@SUP_P", tmp269 ?? DBNull.Value);
            command.Parameters.AddWithValue("@SUP_R", tmp270 ?? DBNull.Value);
            command.Parameters.AddWithValue("@SUP_URB", tmp271 ?? DBNull.Value);
            command.Parameters.AddWithValue("@SUP_URB_P", tmp272 ?? DBNull.Value);
            command.Parameters.AddWithValue("@SUP_URB_R", tmp273 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_VAR_PERC", tmp274 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_VAR_PERC_P", tmp275 ?? DBNull.Value);
            command.Parameters.AddWithValue("@PERC_VAR_PERC_R", tmp276 ?? DBNull.Value);
            command.Parameters.AddWithValue("@VECCH", tmp277 ?? DBNull.Value);
            command.Parameters.AddWithValue("@VECCH_P", tmp278 ?? DBNull.Value);
            command.Parameters.AddWithValue("@VECCH_R", tmp279 ?? DBNull.Value);
            command.Parameters.AddWithValue("@DataImportazione", DateTime.Now);
            command.Parameters.AddWithValue("@ParentWeb_Id", 13029);

            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                Console.WriteLine("Error Generated. Details: " + e.ToString());
            }
            finally
            {
                connection.Close();
            }
        }

        static bool alreadyExists(string comparator)
        {
            Boolean exists = false;
            using (var connection = new SqlConnection(Settings.connection_string))
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT SUP_URB, PERC_POP_ANZ FROM tbl_13029_Stage_Istat";
                    connection.Open();
                    using (var reader = command.ExecuteReader())
                    {
                        var indexOfColumn1 = reader.GetOrdinal("SUP_URB");
                        var indexOfColumn2 = reader.GetOrdinal("PERC_POP_ANZ");


                        while (reader.Read())
                        {
                            string value1 = reader.GetValue(indexOfColumn1).ToString();
                            string value2 = reader.GetValue(indexOfColumn2).ToString();

                            if (value1.Equals("")) value1 = "--";
                            if (value2.Equals("")) value2 = "--";

                            string comparator2 = value1 + value2;
                            if (comparator.Equals(comparator2))
                            {
                                connection.Close();
                                return true;
                            }
                            else
                            {
                                exists = false;
                            }
                        }
                    }
                    connection.Close();
                }
            }
            return exists;
        }

        static int manageExcel(String filename)
        {
            int n_inseriti = 0;
            using (var pck = new ExcelPackage(new FileInfo("../../../Excels/Manuali/" + filename)))
            {
                ExcelWorksheet workSheet = pck.Workbook.Worksheets["Intero territorio nazionale"];
                var start = new ExcelCellAddress(2, 1);
                var end = new ExcelCellAddress(7955, 397);
                for (int row = start.Row; row <= end.Row; row++)
                {

                    object s1 = null;
                    object s2 = null;
                    object s3 = null;
                    object s4 = null;
                    object s5 = null;
                    object s6 = null;
                    object s7 = null;
                    object s8 = null;
                    object s9 = null;
                    object s10 = null;
                    object s11 = null;
                    object s12 = null;
                    object s13 = null;
                    object s14 = null;
                    object s15 = null;
                    object s16 = null;
                    object s17 = null;
                    object s18 = null;
                    object s19 = null;
                    object s20 = null;
                    object s21 = null;
                    object s22 = null;
                    object s23 = null;
                    object s24 = null;
                    object s25 = null;
                    object s26 = null;
                    object s27 = null;
                    object s28 = null;
                    object s29 = null;
                    object s30 = null;
                    object s31 = null;
                    object s32 = null;
                    object s33 = null;
                    object s34 = null;
                    object s35 = null;
                    object s36 = null;
                    object s37 = null;
                    object s38 = null;
                    object s39 = null;
                    object s40 = null;
                    object s41 = null;
                    object s42 = null;
                    object s43 = null;
                    object s44 = null;
                    object s45 = null;
                    object s46 = null;
                    object s47 = null;
                    object s48 = null;
                    object s49 = null;
                    object s50 = null;
                    object s51 = null;
                    object s52 = null;
                    object s53 = null;
                    object s54 = null;
                    object s55 = null;
                    object s56 = null;
                    object s57 = null;
                    object s58 = null;
                    object s59 = null;
                    object s60 = null;
                    object s61 = null;
                    object s62 = null;
                    object s63 = null;
                    object s64 = null;
                    object s65 = null;
                    object s66 = null;
                    object s67 = null;
                    object s68 = null;
                    object s69 = null;
                    object s70 = null;
                    object s71 = null;
                    object s72 = null;
                    object s73 = null;
                    object s74 = null;
                    object s75 = null;
                    object s76 = null;
                    object s77 = null;
                    object s78 = null;
                    object s79 = null;
                    object s80 = null;
                    object s81 = null;
                    object s82 = null;
                    object s83 = null;
                    object s84 = null;
                    object s85 = null;
                    object s86 = null;
                    object s87 = null;
                    object s88 = null;
                    object s89 = null;
                    object s90 = null;
                    object s91 = null;
                    object s92 = null;
                    object s93 = null;
                    object s94 = null;
                    object s95 = null;
                    object s96 = null;
                    object s97 = null;
                    object s98 = null;
                    object s99 = null;
                    object s100 = null;
                    object s101 = null;
                    object s102 = null;
                    object s103 = null;
                    object s104 = null;
                    object s105 = null;
                    object s106 = null;
                    object s107 = null;
                    object s108 = null;
                    object s109 = null;
                    object s110 = null;
                    object s111 = null;
                    object s112 = null;
                    object s113 = null;
                    object s114 = null;
                    object s115 = null;
                    object s116 = null;
                    object s117 = null;
                    object s118 = null;
                    object s119 = null;
                    object s120 = null;
                    object s121 = null;
                    object s122 = null;
                    object s123 = null;
                    object s124 = null;
                    object s125 = null;
                    object s126 = null;
                    object s127 = null;
                    object s128 = null;
                    object s129 = null;
                    object s130 = null;
                    object s131 = null;
                    object s132 = null;
                    object s133 = null;
                    object s134 = null;
                    object s135 = null;
                    object s136 = null;
                    object s137 = null;
                    object s138 = null;
                    object s139 = null;
                    object s140 = null;
                    object s141 = null;
                    object s142 = null;
                    object s143 = null;
                    object s144 = null;
                    object s145 = null;
                    object s146 = null;
                    object s147 = null;
                    object s148 = null;
                    object s149 = null;
                    object s150 = null;
                    object s151 = null;
                    object s152 = null;
                    object s153 = null;
                    object s154 = null;
                    object s155 = null;
                    object s156 = null;
                    object s157 = null;
                    object s158 = null;
                    object s159 = null;
                    object s160 = null;
                    object s161 = null;
                    object s162 = null;
                    object s163 = null;
                    object s164 = null;
                    object s165 = null;
                    object s166 = null;
                    object s167 = null;
                    object s168 = null;
                    object s169 = null;
                    object s170 = null;
                    object s171 = null;
                    object s172 = null;
                    object s173 = null;
                    object s174 = null;
                    object s175 = null;
                    object s176 = null;
                    object s177 = null;
                    object s178 = null;
                    object s179 = null;
                    object s180 = null;
                    object s181 = null;
                    object s182 = null;
                    object s183 = null;
                    object s184 = null;
                    object s185 = null;
                    object s186 = null;
                    object s187 = null;
                    object s188 = null;
                    object s189 = null;
                    object s190 = null;
                    object s191 = null;
                    object s192 = null;
                    object s193 = null;
                    object s194 = null;
                    object s195 = null;
                    object s196 = null;
                    object s197 = null;
                    object s198 = null;
                    object s199 = null;
                    object s200 = null;
                    object s201 = null;
                    object s202 = null;
                    object s203 = null;
                    object s204 = null;
                    object s205 = null;
                    object s206 = null;
                    object s207 = null;
                    object s208 = null;
                    object s209 = null;
                    object s210 = null;
                    object s211 = null;
                    object s212 = null;
                    object s213 = null;
                    object s214 = null;
                    object s215 = null;
                    object s216 = null;
                    object s217 = null;
                    object s218 = null;
                    object s219 = null;
                    object s220 = null;
                    object s221 = null;
                    object s222 = null;
                    object s223 = null;
                    object s224 = null;
                    object s225 = null;
                    object s226 = null;
                    object s227 = null;
                    object s228 = null;
                    object s229 = null;
                    object s230 = null;
                    object s231 = null;
                    object s232 = null;
                    object s233 = null;
                    object s234 = null;
                    object s235 = null;
                    object s236 = null;
                    object s237 = null;
                    object s238 = null;
                    object s239 = null;
                    object s240 = null;
                    object s241 = null;
                    object s242 = null;
                    object s243 = null;
                    object s244 = null;
                    object s245 = null;
                    object s246 = null;
                    object s247 = null;
                    object s248 = null;
                    object s249 = null;
                    object s250 = null;
                    object s251 = null;
                    object s252 = null;
                    object s253 = null;
                    object s254 = null;
                    object s255 = null;
                    object s256 = null;
                    object s257 = null;
                    object s258 = null;
                    object s259 = null;
                    object s260 = null;
                    object s261 = null;
                    object s262 = null;
                    object s263 = null;
                    object s264 = null;
                    object s265 = null;
                    object s266 = null;
                    object s267 = null;
                    object s268 = null;
                    object s269 = null;
                    object s270 = null;
                    object s271 = null;
                    object s272 = null;
                    object s273 = null;
                    object s274 = null;
                    object s275 = null;
                    object s276 = null;
                    object s277 = null;
                    object s278 = null;
                    object s279 = null;

                    //SUP_URB + PERC_POZ_ANZ
                    string comparator = workSheet.Cells[row, 17].Text + workSheet.Cells[row, 43].Text;

                    if (!workSheet.Cells[row, 13].Text.Equals("--")) s1 = Double.Parse(workSheet.Cells[row, 13].Text);
                    if (!workSheet.Cells[row, 206].Text.Equals("--")) s2 = Double.Parse(workSheet.Cells[row, 206].Text);
                    if (!workSheet.Cells[row, 207].Text.Equals("--")) s3 = Double.Parse(workSheet.Cells[row, 207].Text);
                    if (!workSheet.Cells[row, 208].Text.Equals("--")) s4 = Double.Parse(workSheet.Cells[row, 208].Text);
                    if (!workSheet.Cells[row, 209].Text.Equals("--")) s5 = Double.Parse(workSheet.Cells[row, 209].Text);
                    if (!workSheet.Cells[row, 210].Text.Equals("--")) s6 = Double.Parse(workSheet.Cells[row, 210].Text);
                    if (!workSheet.Cells[row, 211].Text.Equals("--")) s7 = Double.Parse(workSheet.Cells[row, 211].Text);
                    if (!workSheet.Cells[row, 227].Text.Equals("--")) s8 = Double.Parse(workSheet.Cells[row, 227].Text);
                    if (!workSheet.Cells[row, 228].Text.Equals("--")) s9 = Double.Parse(workSheet.Cells[row, 228].Text);
                    if (!workSheet.Cells[row, 229].Text.Equals("--")) s10 = Double.Parse(workSheet.Cells[row, 229].Text);
                    if (!workSheet.Cells[row, 35].Text.Equals("--")) s11 = Double.Parse(workSheet.Cells[row, 35].Text);
                    if (!workSheet.Cells[row, 36].Text.Equals("--")) s12 = workSheet.Cells[row, 36].Text;
                    if (!workSheet.Cells[row, 37].Text.Equals("--")) s13 = Double.Parse(workSheet.Cells[row, 37].Text);
                    if (!workSheet.Cells[row, 38].Text.Equals("--")) s14 = Double.Parse(workSheet.Cells[row, 38].Text);
                    if (!workSheet.Cells[row, 39].Text.Equals("--")) s15 = Double.Parse(workSheet.Cells[row, 39].Text);//DENSPOP_R
                    if (!workSheet.Cells[row, 344].Text.Equals("--")) s16 = Double.Parse(workSheet.Cells[row, 344].Text);
                    if (!workSheet.Cells[row, 345].Text.Equals("--")) s17 = Double.Parse(workSheet.Cells[row, 345].Text);
                    if (!workSheet.Cells[row, 346].Text.Equals("--")) s18 = Double.Parse(workSheet.Cells[row, 346].Text);
                    if (!workSheet.Cells[row, 347].Text.Equals("--")) s19 = Double.Parse(workSheet.Cells[row, 347].Text);
                    if (!workSheet.Cells[row, 348].Text.Equals("--")) s20 = Double.Parse(workSheet.Cells[row, 348].Text);
                    if (!workSheet.Cells[row, 349].Text.Equals("--")) s21 = Double.Parse(workSheet.Cells[row, 349].Text);
                    if (!workSheet.Cells[row, 350].Text.Equals("--")) s22 = Double.Parse(workSheet.Cells[row, 350].Text);
                    if (!workSheet.Cells[row, 351].Text.Equals("--")) s23 = Double.Parse(workSheet.Cells[row, 351].Text);
                    if (!workSheet.Cells[row, 352].Text.Equals("--")) s24 = Double.Parse(workSheet.Cells[row, 352].Text);//PERC_ECP3_16_R
                    if (!workSheet.Cells[row, 335].Text.Equals("--")) s25 = Double.Parse(workSheet.Cells[row, 335].Text);//PERC_EC8_12
                    if (!workSheet.Cells[row, 337].Text.Equals("--")) s26 = Double.Parse(workSheet.Cells[row, 337].Text);//PERC_EC16
                    if (!workSheet.Cells[row, 336].Text.Equals("--")) s27 = Double.Parse(workSheet.Cells[row, 336].Text);//PERC_EC13_15
                    if (!workSheet.Cells[row, 338].Text.Equals("--")) s28 = Double.Parse(workSheet.Cells[row, 338].Text);//PERC_EC8_12_P
                    if (!workSheet.Cells[row, 339].Text.Equals("--")) s29 = Double.Parse(workSheet.Cells[row, 339].Text);//PERC_EC13_15_P
                    if (!workSheet.Cells[row, 340].Text.Equals("--")) s30 = Double.Parse(workSheet.Cells[row, 340].Text);//PERC_EC16_P
                    if (!workSheet.Cells[row, 341].Text.Equals("--")) s31 = Double.Parse(workSheet.Cells[row, 341].Text);//PERC_EC8_12_R
                    if (!workSheet.Cells[row, 342].Text.Equals("--")) s32 = Double.Parse(workSheet.Cells[row, 342].Text);//PERC_EC13_15_R
                    if (!workSheet.Cells[row, 343].Text.Equals("--")) s33 = Double.Parse(workSheet.Cells[row, 343].Text);//PERC_EC16_R
                    if (!workSheet.Cells[row, 269].Text.Equals("--")) s34 = Double.Parse(workSheet.Cells[row, 269].Text);
                    if (!workSheet.Cells[row, 270].Text.Equals("--")) s35 = Double.Parse(workSheet.Cells[row, 270].Text);
                    if (!workSheet.Cells[row, 271].Text.Equals("--")) s36 = Double.Parse(workSheet.Cells[row, 271].Text);
                    if (!workSheet.Cells[row, 272].Text.Equals("--")) s37 = Double.Parse(workSheet.Cells[row, 272].Text);
                    if (!workSheet.Cells[row, 273].Text.Equals("--")) s38 = Double.Parse(workSheet.Cells[row, 273].Text);
                    if (!workSheet.Cells[row, 274].Text.Equals("--")) s39 = Double.Parse(workSheet.Cells[row, 274].Text);
                    if (!workSheet.Cells[row, 275].Text.Equals("--")) s40 = Double.Parse(workSheet.Cells[row, 275].Text);
                    if (!workSheet.Cells[row, 276].Text.Equals("--")) s41 = Double.Parse(workSheet.Cells[row, 276].Text);
                    if (!workSheet.Cells[row, 277].Text.Equals("--")) s42 = Double.Parse(workSheet.Cells[row, 277].Text);
                    if (!workSheet.Cells[row, 278].Text.Equals("--")) s43 = Double.Parse(workSheet.Cells[row, 278].Text);
                    if (!workSheet.Cells[row, 279].Text.Equals("--")) s44 = Double.Parse(workSheet.Cells[row, 279].Text);
                    if (!workSheet.Cells[row, 280].Text.Equals("--")) s45 = Double.Parse(workSheet.Cells[row, 280].Text);//PERC_EMC_P_R
                    if (!workSheet.Cells[row, 233].Text.Equals("--")) s46 = Double.Parse(workSheet.Cells[row, 233].Text);
                    if (!workSheet.Cells[row, 234].Text.Equals("--")) s47 = Double.Parse(workSheet.Cells[row, 234].Text);
                    if (!workSheet.Cells[row, 235].Text.Equals("--")) s48 = Double.Parse(workSheet.Cells[row, 235].Text);
                    if (!workSheet.Cells[row, 236].Text.Equals("--")) s49 = Double.Parse(workSheet.Cells[row, 236].Text);
                    if (!workSheet.Cells[row, 237].Text.Equals("--")) s50 = Double.Parse(workSheet.Cells[row, 237].Text);
                    if (!workSheet.Cells[row, 238].Text.Equals("--")) s51 = Double.Parse(workSheet.Cells[row, 238].Text);
                    if (!workSheet.Cells[row, 239].Text.Equals("--")) s52 = Double.Parse(workSheet.Cells[row, 239].Text);
                    if (!workSheet.Cells[row, 240].Text.Equals("--")) s53 = Double.Parse(workSheet.Cells[row, 240].Text);
                    if (!workSheet.Cells[row, 241].Text.Equals("--")) s54 = Double.Parse(workSheet.Cells[row, 241].Text);
                    if (!workSheet.Cells[row, 242].Text.Equals("--")) s55 = Double.Parse(workSheet.Cells[row, 242].Text);
                    if (!workSheet.Cells[row, 243].Text.Equals("--")) s56 = Double.Parse(workSheet.Cells[row, 243].Text);
                    if (!workSheet.Cells[row, 244].Text.Equals("--")) s57 = Double.Parse(workSheet.Cells[row, 244].Text);
                    if (!workSheet.Cells[row, 245].Text.Equals("--")) s58 = Double.Parse(workSheet.Cells[row, 245].Text);
                    if (!workSheet.Cells[row, 246].Text.Equals("--")) s59 = Double.Parse(workSheet.Cells[row, 246].Text);
                    if (!workSheet.Cells[row, 247].Text.Equals("--")) s60 = Double.Parse(workSheet.Cells[row, 247].Text);
                    if (!workSheet.Cells[row, 248].Text.Equals("--")) s61 = Double.Parse(workSheet.Cells[row, 248].Text);
                    if (!workSheet.Cells[row, 249].Text.Equals("--")) s62 = Double.Parse(workSheet.Cells[row, 249].Text);
                    if (!workSheet.Cells[row, 250].Text.Equals("--")) s63 = Double.Parse(workSheet.Cells[row, 250].Text);
                    if (!workSheet.Cells[row, 251].Text.Equals("--")) s64 = Double.Parse(workSheet.Cells[row, 251].Text);
                    if (!workSheet.Cells[row, 252].Text.Equals("--")) s65 = Double.Parse(workSheet.Cells[row, 252].Text);
                    if (!workSheet.Cells[row, 253].Text.Equals("--")) s66 = Double.Parse(workSheet.Cells[row, 253].Text);
                    if (!workSheet.Cells[row, 254].Text.Equals("--")) s67 = Double.Parse(workSheet.Cells[row, 254].Text);
                    if (!workSheet.Cells[row, 255].Text.Equals("--")) s68 = Double.Parse(workSheet.Cells[row, 255].Text);
                    if (!workSheet.Cells[row, 256].Text.Equals("--")) s69 = Double.Parse(workSheet.Cells[row, 256].Text);
                    if (!workSheet.Cells[row, 257].Text.Equals("--")) s70 = Double.Parse(workSheet.Cells[row, 257].Text);
                    if (!workSheet.Cells[row, 258].Text.Equals("--")) s71 = Double.Parse(workSheet.Cells[row, 258].Text);
                    if (!workSheet.Cells[row, 259].Text.Equals("--")) s72 = Double.Parse(workSheet.Cells[row, 259].Text);
                    if (!workSheet.Cells[row, 260].Text.Equals("--")) s73 = Double.Parse(workSheet.Cells[row, 260].Text);
                    if (!workSheet.Cells[row, 261].Text.Equals("--")) s74 = Double.Parse(workSheet.Cells[row, 261].Text);
                    if (!workSheet.Cells[row, 262].Text.Equals("--")) s75 = Double.Parse(workSheet.Cells[row, 262].Text);
                    if (!workSheet.Cells[row, 263].Text.Equals("--")) s76 = Double.Parse(workSheet.Cells[row, 263].Text);
                    if (!workSheet.Cells[row, 264].Text.Equals("--")) s77 = Double.Parse(workSheet.Cells[row, 264].Text);
                    if (!workSheet.Cells[row, 265].Text.Equals("--")) s78 = Double.Parse(workSheet.Cells[row, 265].Text);
                    if (!workSheet.Cells[row, 266].Text.Equals("--")) s79 = Double.Parse(workSheet.Cells[row, 266].Text);
                    if (!workSheet.Cells[row, 267].Text.Equals("--")) s80 = Double.Parse(workSheet.Cells[row, 267].Text);
                    if (!workSheet.Cells[row, 268].Text.Equals("--")) s81 = Double.Parse(workSheet.Cells[row, 268].Text);//PERC_EMP3_R
                    if (!workSheet.Cells[row, 332].Text.Equals("--")) s82 = Double.Parse(workSheet.Cells[row, 332].Text);
                    if (!workSheet.Cells[row, 333].Text.Equals("--")) s83 = Double.Parse(workSheet.Cells[row, 333].Text);
                    if (!workSheet.Cells[row, 334].Text.Equals("--")) s84 = Double.Parse(workSheet.Cells[row, 334].Text);//PERC_EP3C_R
                    if (!workSheet.Cells[row, 311].Text.Equals("--")) s85 = Double.Parse(workSheet.Cells[row, 311].Text);
                    if (!workSheet.Cells[row, 312].Text.Equals("--")) s86 = Double.Parse(workSheet.Cells[row, 312].Text);
                    if (!workSheet.Cells[row, 313].Text.Equals("--")) s87 = Double.Parse(workSheet.Cells[row, 313].Text);
                    if (!workSheet.Cells[row, 314].Text.Equals("--")) s88 = Double.Parse(workSheet.Cells[row, 314].Text);
                    if (!workSheet.Cells[row, 315].Text.Equals("--")) s89 = Double.Parse(workSheet.Cells[row, 315].Text);
                    if (!workSheet.Cells[row, 316].Text.Equals("--")) s90 = Double.Parse(workSheet.Cells[row, 316].Text);
                    if (!workSheet.Cells[row, 317].Text.Equals("--")) s91 = Double.Parse(workSheet.Cells[row, 317].Text);
                    if (!workSheet.Cells[row, 318].Text.Equals("--")) s92 = Double.Parse(workSheet.Cells[row, 318].Text);
                    if (!workSheet.Cells[row, 319].Text.Equals("--")) s93 = Double.Parse(workSheet.Cells[row, 319].Text);
                    if (!workSheet.Cells[row, 320].Text.Equals("--")) s94 = Double.Parse(workSheet.Cells[row, 320].Text);
                    if (!workSheet.Cells[row, 321].Text.Equals("--")) s95 = Double.Parse(workSheet.Cells[row, 321].Text);
                    if (!workSheet.Cells[row, 322].Text.Equals("--")) s96 = Double.Parse(workSheet.Cells[row, 322].Text);//PERC_EP3C4_R
                    if (!workSheet.Cells[row, 281].Text.Equals("--")) s97 = Double.Parse(workSheet.Cells[row, 281].Text);
                    if (!workSheet.Cells[row, 282].Text.Equals("--")) s98 = Double.Parse(workSheet.Cells[row, 282].Text);
                    if (!workSheet.Cells[row, 283].Text.Equals("--")) s99 = Double.Parse(workSheet.Cells[row, 283].Text);
                    if (!workSheet.Cells[row, 284].Text.Equals("--")) s100 = Double.Parse(workSheet.Cells[row, 284].Text);
                    if (!workSheet.Cells[row, 285].Text.Equals("--")) s101 = Double.Parse(workSheet.Cells[row, 285].Text);
                    if (!workSheet.Cells[row, 286].Text.Equals("--")) s102 = Double.Parse(workSheet.Cells[row, 286].Text);
                    if (!workSheet.Cells[row, 287].Text.Equals("--")) s103 = Double.Parse(workSheet.Cells[row, 287].Text);
                    if (!workSheet.Cells[row, 288].Text.Equals("--")) s104 = Double.Parse(workSheet.Cells[row, 288].Text);
                    if (!workSheet.Cells[row, 289].Text.Equals("--")) s105 = Double.Parse(workSheet.Cells[row, 289].Text);
                    if (!workSheet.Cells[row, 290].Text.Equals("--")) s106 = Double.Parse(workSheet.Cells[row, 290].Text);
                    if (!workSheet.Cells[row, 291].Text.Equals("--")) s107 = Double.Parse(workSheet.Cells[row, 291].Text);
                    if (!workSheet.Cells[row, 292].Text.Equals("--")) s108 = Double.Parse(workSheet.Cells[row, 292].Text);
                    if (!workSheet.Cells[row, 293].Text.Equals("--")) s109 = Double.Parse(workSheet.Cells[row, 293].Text);
                    if (!workSheet.Cells[row, 294].Text.Equals("--")) s110 = Double.Parse(workSheet.Cells[row, 294].Text);
                    if (!workSheet.Cells[row, 295].Text.Equals("--")) s111 = Double.Parse(workSheet.Cells[row, 295].Text);
                    if (!workSheet.Cells[row, 296].Text.Equals("--")) s112 = Double.Parse(workSheet.Cells[row, 296].Text);
                    if (!workSheet.Cells[row, 297].Text.Equals("--")) s113 = Double.Parse(workSheet.Cells[row, 297].Text);
                    if (!workSheet.Cells[row, 298].Text.Equals("--")) s114 = Double.Parse(workSheet.Cells[row, 298].Text);
                    if (!workSheet.Cells[row, 299].Text.Equals("--")) s115 = Double.Parse(workSheet.Cells[row, 299].Text);
                    if (!workSheet.Cells[row, 300].Text.Equals("--")) s116 = Double.Parse(workSheet.Cells[row, 300].Text);
                    if (!workSheet.Cells[row, 301].Text.Equals("--")) s117 = Double.Parse(workSheet.Cells[row, 301].Text);
                    if (!workSheet.Cells[row, 302].Text.Equals("--")) s118 = Double.Parse(workSheet.Cells[row, 302].Text);
                    if (!workSheet.Cells[row, 303].Text.Equals("--")) s119 = Double.Parse(workSheet.Cells[row, 303].Text);
                    if (!workSheet.Cells[row, 304].Text.Equals("--")) s120 = Double.Parse(workSheet.Cells[row, 304].Text);
                    if (!workSheet.Cells[row, 305].Text.Equals("--")) s121 = Double.Parse(workSheet.Cells[row, 305].Text);
                    if (!workSheet.Cells[row, 306].Text.Equals("--")) s122 = Double.Parse(workSheet.Cells[row, 306].Text);
                    if (!workSheet.Cells[row, 307].Text.Equals("--")) s123 = Double.Parse(workSheet.Cells[row, 307].Text);
                    if (!workSheet.Cells[row, 308].Text.Equals("--")) s124 = Double.Parse(workSheet.Cells[row, 308].Text);
                    if (!workSheet.Cells[row, 309].Text.Equals("--")) s125 = Double.Parse(workSheet.Cells[row, 309].Text);
                    if (!workSheet.Cells[row, 310].Text.Equals("--")) s126 = Double.Parse(workSheet.Cells[row, 310].Text);//PERC_EP3M_R
                    if (!workSheet.Cells[row, 325].Text.Equals("--")) s127 = Double.Parse(workSheet.Cells[row, 325].Text);
                    if (!workSheet.Cells[row, 328].Text.Equals("--")) s128 = Double.Parse(workSheet.Cells[row, 328].Text);
                    if (!workSheet.Cells[row, 331].Text.Equals("--")) s129 = Double.Parse(workSheet.Cells[row, 331].Text);//PERC_ERA1_R
                    if (!workSheet.Cells[row, 324].Text.Equals("--")) s130 = Double.Parse(workSheet.Cells[row, 324].Text);
                    if (!workSheet.Cells[row, 327].Text.Equals("--")) s131 = Double.Parse(workSheet.Cells[row, 327].Text);
                    if (!workSheet.Cells[row, 330].Text.Equals("--")) s132 = Double.Parse(workSheet.Cells[row, 330].Text);//PERC_ERC1_R
                    if (!workSheet.Cells[row, 395].Text.Equals("--")) s133 = Double.Parse(workSheet.Cells[row, 395].Text);
                    if (!workSheet.Cells[row, 393].Text.Equals("--")) s134 = Double.Parse(workSheet.Cells[row, 393].Text);
                    if (!workSheet.Cells[row, 394].Text.Equals("--")) s135 = workSheet.Cells[row, 394].Text;
                    if (!workSheet.Cells[row, 397].Text.Equals("--")) s136 = Double.Parse(workSheet.Cells[row, 397].Text);
                    if (!workSheet.Cells[row, 396].Text.Equals("--")) s137 = Double.Parse(workSheet.Cells[row, 396].Text);//ERESPP
                    if (!workSheet.Cells[row, 353].Text.Equals("--")) s138 = Double.Parse(workSheet.Cells[row, 353].Text);//ERE8
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s139 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE8_P
                    if (!workSheet.Cells[row, 371].Text.Equals("--")) s140 = Double.Parse(workSheet.Cells[row, 371].Text);//ERE8_R
                    if (!workSheet.Cells[row, 354].Text.Equals("--")) s141 = Double.Parse(workSheet.Cells[row, 354].Text);//ERE9
                    if (!workSheet.Cells[row, 363].Text.Equals("--")) s142 = Double.Parse(workSheet.Cells[row, 363].Text);//ERE9_P
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s143 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE9_R
                    if (!workSheet.Cells[row, 355].Text.Equals("--")) s144 = Double.Parse(workSheet.Cells[row, 355].Text);//ERE10
                    if (!workSheet.Cells[row, 364].Text.Equals("--")) s145 = Double.Parse(workSheet.Cells[row, 364].Text);//ERE10_P
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s146 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE10_R
                    if (!workSheet.Cells[row, 356].Text.Equals("--")) s147 = Double.Parse(workSheet.Cells[row, 356].Text);//ERE11
                    if (!workSheet.Cells[row, 365].Text.Equals("--")) s148 = Double.Parse(workSheet.Cells[row, 365].Text);//ERE11_P
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s149 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE11_R
                    if (!workSheet.Cells[row, 357].Text.Equals("--")) s150 = Double.Parse(workSheet.Cells[row, 357].Text);//ERE12
                    if (!workSheet.Cells[row, 366].Text.Equals("--")) s151 = Double.Parse(workSheet.Cells[row, 366].Text);//ERE12_P
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s152 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE12_R
                    if (!workSheet.Cells[row, 358].Text.Equals("--")) s153 = Double.Parse(workSheet.Cells[row, 358].Text);//ERE13
                    if (!workSheet.Cells[row, 367].Text.Equals("--")) s154 = Double.Parse(workSheet.Cells[row, 367].Text);//ERE13_P
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s155 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE13_R
                    if (!workSheet.Cells[row, 359].Text.Equals("--")) s156 = Double.Parse(workSheet.Cells[row, 359].Text);//ERE14
                    if (!workSheet.Cells[row, 368].Text.Equals("--")) s157 = Double.Parse(workSheet.Cells[row, 368].Text);//ERE14_P
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s158 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE14_R
                    if (!workSheet.Cells[row, 360].Text.Equals("--")) s159 = Double.Parse(workSheet.Cells[row, 360].Text);//ERE15
                    if (!workSheet.Cells[row, 369].Text.Equals("--")) s160 = Double.Parse(workSheet.Cells[row, 369].Text);//ERE15_P
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s161 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE15_R
                    if (!workSheet.Cells[row, 361].Text.Equals("--")) s162 = Double.Parse(workSheet.Cells[row, 361].Text);//ERE16
                    if (!workSheet.Cells[row, 370].Text.Equals("--")) s163 = Double.Parse(workSheet.Cells[row, 370].Text);//ERE16_P
                    if (!workSheet.Cells[row, 362].Text.Equals("--")) s164 = Double.Parse(workSheet.Cells[row, 362].Text);//ERE16_R
                    if (!workSheet.Cells[row, 323].Text.Equals("--")) s165 = Double.Parse(workSheet.Cells[row, 323].Text);//PERC_ERM1
                    if (!workSheet.Cells[row, 326].Text.Equals("--")) s166 = Double.Parse(workSheet.Cells[row, 326].Text);//PERC_ERM1_P
                    if (!workSheet.Cells[row, 329].Text.Equals("--")) s167 = Double.Parse(workSheet.Cells[row, 329].Text);//PERC_ERM1_R
                    if (!workSheet.Cells[row, 58].Text.Equals("--")) s168 = Double.Parse(workSheet.Cells[row, 58].Text);//PERC_ETA_Q1
                    if (!workSheet.Cells[row, 65].Text.Equals("--")) s169 = Double.Parse(workSheet.Cells[row, 65].Text);//PERC_ETA_Q1_P
                    if (!workSheet.Cells[row, 66].Text.Equals("--")) s170 = Double.Parse(workSheet.Cells[row, 66].Text);//PERC_ETA_Q1_R
                    if (!workSheet.Cells[row, 59].Text.Equals("--")) s171 = Double.Parse(workSheet.Cells[row, 59].Text);//PERC_ETA_Q2
                    if (!workSheet.Cells[row, 64].Text.Equals("--")) s172 = Double.Parse(workSheet.Cells[row, 64].Text);//PERC_ETA_Q2_P
                    if (!workSheet.Cells[row, 67].Text.Equals("--")) s173 = Double.Parse(workSheet.Cells[row, 67].Text);//PERC_ETA_Q2_R
                    if (!workSheet.Cells[row, 60].Text.Equals("--")) s174 = Double.Parse(workSheet.Cells[row, 60].Text);//PERC_ETA_Q3
                    if (!workSheet.Cells[row, 63].Text.Equals("--")) s175 = Double.Parse(workSheet.Cells[row, 63].Text);//PERC_ETA_Q3_P
                    if (!workSheet.Cells[row, 68].Text.Equals("--")) s176 = Double.Parse(workSheet.Cells[row, 68].Text);//PERC_ETA_Q3_R
                    if (!workSheet.Cells[row, 61].Text.Equals("--")) s177 = Double.Parse(workSheet.Cells[row, 61].Text);//PERC_ETA_Q4
                    if (!workSheet.Cells[row, 62].Text.Equals("--")) s178 = Double.Parse(workSheet.Cells[row, 62].Text);//PERC_ETA_Q4_P
                    if (!workSheet.Cells[row, 69].Text.Equals("--")) s179 = Double.Parse(workSheet.Cells[row, 69].Text);//PERC_ETA_Q4_R
                    if (!workSheet.Cells[row, 143].Text.Equals("--")) s180 = Double.Parse(workSheet.Cells[row, 143].Text);//E3
                    if (!workSheet.Cells[row, 144].Text.Equals("--")) s181 = Double.Parse(workSheet.Cells[row, 144].Text);//E3_P
                    if (!workSheet.Cells[row, 145].Text.Equals("--")) s182 = Double.Parse(workSheet.Cells[row, 145].Text);//E3_R
                    if (!workSheet.Cells[row, 146].Text.Equals("--")) s183 = Double.Parse(workSheet.Cells[row, 146].Text);//E5
                    if (!workSheet.Cells[row, 151].Text.Equals("--")) s184 = Double.Parse(workSheet.Cells[row, 151].Text);//E5_P
                    if (!workSheet.Cells[row, 152].Text.Equals("--")) s185 = Double.Parse(workSheet.Cells[row, 152].Text);//E5_R
                    if (!workSheet.Cells[row, 147].Text.Equals("--")) s186 = Double.Parse(workSheet.Cells[row, 147].Text);//E6
                    if (!workSheet.Cells[row, 149].Text.Equals("--")) s187 = Double.Parse(workSheet.Cells[row, 149].Text);//E6_P
                    if (!workSheet.Cells[row, 150].Text.Equals("--")) s188 = Double.Parse(workSheet.Cells[row, 150].Text);//E6_R
                    if (!workSheet.Cells[row, 148].Text.Equals("--")) s189 = Double.Parse(workSheet.Cells[row, 148].Text);//E7
                    if (!workSheet.Cells[row, 153].Text.Equals("--")) s190 = Double.Parse(workSheet.Cells[row, 153].Text);//E7_P
                    if (!workSheet.Cells[row, 154].Text.Equals("--")) s191 = Double.Parse(workSheet.Cells[row, 154].Text);//E7_R
                    if (!workSheet.Cells[row, 155].Text.Equals("--")) s192 = Double.Parse(workSheet.Cells[row, 155].Text);//E8
                    if (!workSheet.Cells[row, 166].Text.Equals("--")) s193 = Double.Parse(workSheet.Cells[row, 166].Text);//E8_P
                    if (!workSheet.Cells[row, 167].Text.Equals("--")) s194 = Double.Parse(workSheet.Cells[row, 167].Text);//E8_R
                    if (!workSheet.Cells[row, 156].Text.Equals("--")) s195 = Double.Parse(workSheet.Cells[row, 156].Text);//E9
                    if (!workSheet.Cells[row, 168].Text.Equals("--")) s196 = Double.Parse(workSheet.Cells[row, 168].Text);//E9_P
                    if (!workSheet.Cells[row, 169].Text.Equals("--")) s197 = Double.Parse(workSheet.Cells[row, 169].Text);//E9_R
                    if (!workSheet.Cells[row, 157].Text.Equals("--")) s198 = Double.Parse(workSheet.Cells[row, 157].Text);//E10
                    if (!workSheet.Cells[row, 170].Text.Equals("--")) s199 = Double.Parse(workSheet.Cells[row, 170].Text);//E10_P
                    if (!workSheet.Cells[row, 171].Text.Equals("--")) s200 = Double.Parse(workSheet.Cells[row, 171].Text);//E10_R
                    if (!workSheet.Cells[row, 158].Text.Equals("--")) s201 = Double.Parse(workSheet.Cells[row, 158].Text);//E11
                    if (!workSheet.Cells[row, 172].Text.Equals("--")) s202 = Double.Parse(workSheet.Cells[row, 172].Text);//E11_P
                    if (!workSheet.Cells[row, 173].Text.Equals("--")) s203 = Double.Parse(workSheet.Cells[row, 173].Text);//E11_R
                    if (!workSheet.Cells[row, 159].Text.Equals("--")) s204 = Double.Parse(workSheet.Cells[row, 159].Text);//E12
                    if (!workSheet.Cells[row, 174].Text.Equals("--")) s205 = Double.Parse(workSheet.Cells[row, 174].Text);//E12_P
                    if (!workSheet.Cells[row, 175].Text.Equals("--")) s206 = Double.Parse(workSheet.Cells[row, 175].Text);//E12_R
                    if (!workSheet.Cells[row, 160].Text.Equals("--")) s207 = Double.Parse(workSheet.Cells[row, 160].Text);//E13
                    if (!workSheet.Cells[row, 176].Text.Equals("--")) s208 = Double.Parse(workSheet.Cells[row, 176].Text);//E13_P
                    if (!workSheet.Cells[row, 177].Text.Equals("--")) s209 = Double.Parse(workSheet.Cells[row, 177].Text);//E13_R
                    if (!workSheet.Cells[row, 161].Text.Equals("--")) s210 = Double.Parse(workSheet.Cells[row, 161].Text);//E14
                    if (!workSheet.Cells[row, 178].Text.Equals("--")) s211 = Double.Parse(workSheet.Cells[row, 178].Text);//E14_P
                    if (!workSheet.Cells[row, 179].Text.Equals("--")) s212 = Double.Parse(workSheet.Cells[row, 179].Text);//E14_R
                    if (!workSheet.Cells[row, 162].Text.Equals("--")) s213 = Double.Parse(workSheet.Cells[row, 162].Text);//E15
                    if (!workSheet.Cells[row, 180].Text.Equals("--")) s214 = Double.Parse(workSheet.Cells[row, 180].Text);//E15_P
                    if (!workSheet.Cells[row, 181].Text.Equals("--")) s215 = Double.Parse(workSheet.Cells[row, 181].Text);//E15_R
                    if (!workSheet.Cells[row, 163].Text.Equals("--")) s216 = Double.Parse(workSheet.Cells[row, 163].Text);//E16
                    if (!workSheet.Cells[row, 164].Text.Equals("--")) s217 = Double.Parse(workSheet.Cells[row, 164].Text);//E16_P
                    if (!workSheet.Cells[row, 165].Text.Equals("--")) s218 = Double.Parse(workSheet.Cells[row, 165].Text);//E16_R
                    if (!workSheet.Cells[row, 191].Text.Equals("--")) s219 = Double.Parse(workSheet.Cells[row, 191].Text);//E17
                    if (!workSheet.Cells[row, 195].Text.Equals("--")) s220 = Double.Parse(workSheet.Cells[row, 195].Text);//E17_P
                    if (!workSheet.Cells[row, 196].Text.Equals("--")) s221 = Double.Parse(workSheet.Cells[row, 196].Text);//E17_R
                    if (!workSheet.Cells[row, 192].Text.Equals("--")) s222 = Double.Parse(workSheet.Cells[row, 192].Text);//E18
                    if (!workSheet.Cells[row, 197].Text.Equals("--")) s223 = Double.Parse(workSheet.Cells[row, 197].Text);//E18_P
                    if (!workSheet.Cells[row, 198].Text.Equals("--")) s224 = Double.Parse(workSheet.Cells[row, 198].Text);//E18_R
                    if (!workSheet.Cells[row, 193].Text.Equals("--")) s225 = Double.Parse(workSheet.Cells[row, 193].Text);//E19
                    if (!workSheet.Cells[row, 201].Text.Equals("--")) s226 = Double.Parse(workSheet.Cells[row, 201].Text);//E19_P
                    if (!workSheet.Cells[row, 202].Text.Equals("--")) s227 = Double.Parse(workSheet.Cells[row, 202].Text);//E19_R
                    if (!workSheet.Cells[row, 194].Text.Equals("--")) s228 = Double.Parse(workSheet.Cells[row, 194].Text);//E20
                    if (!workSheet.Cells[row, 199].Text.Equals("--")) s229 = Double.Parse(workSheet.Cells[row, 199].Text);//E20_P
                    if (!workSheet.Cells[row, 200].Text.Equals("--")) s230 = Double.Parse(workSheet.Cells[row, 200].Text);//E20_R
                    if (!workSheet.Cells[row, 215].Text.Equals("--")) s231 = Double.Parse(workSheet.Cells[row, 215].Text);//E28
                    if (!workSheet.Cells[row, 216].Text.Equals("--")) s232 = Double.Parse(workSheet.Cells[row, 216].Text);//E28_P
                    if (!workSheet.Cells[row, 217].Text.Equals("--")) s233 = Double.Parse(workSheet.Cells[row, 217].Text);//E28_R
                    if (!workSheet.Cells[row, 218].Text.Equals("--")) s234 = Double.Parse(workSheet.Cells[row, 218].Text);//E29
                    if (!workSheet.Cells[row, 219].Text.Equals("--")) s235 = Double.Parse(workSheet.Cells[row, 219].Text);//E29_P
                    if (!workSheet.Cells[row, 220].Text.Equals("--")) s236 = Double.Parse(workSheet.Cells[row, 220].Text);//E29_R
                    if (!workSheet.Cells[row, 221].Text.Equals("--")) s237 = Double.Parse(workSheet.Cells[row, 221].Text);//E30
                    if (!workSheet.Cells[row, 222].Text.Equals("--")) s238 = Double.Parse(workSheet.Cells[row, 222].Text);//E30_P
                    if (!workSheet.Cells[row, 223].Text.Equals("--")) s239 = Double.Parse(workSheet.Cells[row, 223].Text);//E30_R
                    if (!workSheet.Cells[row, 224].Text.Equals("--")) s240 = Double.Parse(workSheet.Cells[row, 224].Text);//E31
                    if (!workSheet.Cells[row, 225].Text.Equals("--")) s241 = Double.Parse(workSheet.Cells[row, 225].Text);//E31_P
                    if (!workSheet.Cells[row, 226].Text.Equals("--")) s242 = Double.Parse(workSheet.Cells[row, 226].Text);//E31_R
                    if (!workSheet.Cells[row, 52].Text.Equals("--")) s243 = Double.Parse(workSheet.Cells[row, 52].Text);//FAM_2011
                    if (!workSheet.Cells[row, 53].Text.Equals("--")) s244 = Double.Parse(workSheet.Cells[row, 53].Text);//FAM_2011_p
                    if (!workSheet.Cells[row, 54].Text.Equals("--")) s245 = Double.Parse(workSheet.Cells[row, 54].Text);//FAM_2011_R
                    if (!workSheet.Cells[row, 77].Text.Equals("--")) s246 = Double.Parse(workSheet.Cells[row, 77].Text);//FAM_2018
                    if (!workSheet.Cells[row, 78].Text.Equals("--")) s247 = Double.Parse(workSheet.Cells[row, 78].Text);//FAM_2018_P
                    if (!workSheet.Cells[row, 79].Text.Equals("--")) s248 = Double.Parse(workSheet.Cells[row, 79].Text);//FAM_2018_R
                    if (!workSheet.Cells[row, 70].Text.Equals("--")) s249 = Double.Parse(workSheet.Cells[row, 70].Text);//IDEM
                    if (!workSheet.Cells[row, 71].Text.Equals("--")) s250 = Double.Parse(workSheet.Cells[row, 71].Text);//IDEM_P
                    if (!workSheet.Cells[row, 72].Text.Equals("--")) s251 = Double.Parse(workSheet.Cells[row, 72].Text);//IDEM_R
                    if (!workSheet.Cells[row, 46].Text.Equals("--")) s252 = Double.Parse(workSheet.Cells[row, 46].Text);//PERC_IND_DIP_STR
                    if (!workSheet.Cells[row, 47].Text.Equals("--")) s253 = Double.Parse(workSheet.Cells[row, 47].Text);//PERC_IND_DIP_STR_P
                    if (!workSheet.Cells[row, 48].Text.Equals("--")) s254 = Double.Parse(workSheet.Cells[row, 48].Text);//PERC_IND_DIP_STR_R
                    if (!workSheet.Cells[row, 73].Text.Equals("--")) s255 = Double.Parse(workSheet.Cells[row, 73].Text);//PERC_IVSM
                    if (!workSheet.Cells[row, 9].Text.Equals("--")) s256 = Double.Parse(workSheet.Cells[row, 9].Text);//LIT
                    if (!workSheet.Cells[row, 10].Text.Equals("--")) s257 = workSheet.Cells[row, 10].Text;//LIT_TXT
                    if (!workSheet.Cells[row, 43].Text.Equals("--")) s258 = Double.Parse(workSheet.Cells[row, 43].Text);//PERC_POP_ANZ
                    if (!workSheet.Cells[row, 44].Text.Equals("--")) s259 = Double.Parse(workSheet.Cells[row, 44].Text);//PERC_POP_ANZ_P
                    if (!workSheet.Cells[row, 45].Text.Equals("--")) s260 = Double.Parse(workSheet.Cells[row, 45].Text);//PERC_POP_ANZ_R
                    if (!workSheet.Cells[row, 49].Text.Equals("--")) s261 = Double.Parse(workSheet.Cells[row, 49].Text);//POP_2011
                    if (!workSheet.Cells[row, 50].Text.Equals("--")) s262 = Double.Parse(workSheet.Cells[row, 50].Text);//POP_2011_P
                    if (!workSheet.Cells[row, 51].Text.Equals("--")) s263 = Double.Parse(workSheet.Cells[row, 51].Text);//POP_2011_R
                    if (!workSheet.Cells[row, 74].Text.Equals("--")) s264 = Double.Parse(workSheet.Cells[row, 74].Text);//POP_2018
                    if (!workSheet.Cells[row, 75].Text.Equals("--")) s265 = Double.Parse(workSheet.Cells[row, 75].Text);//POP_2018_P
                    if (!workSheet.Cells[row, 76].Text.Equals("--")) s266 = Double.Parse(workSheet.Cells[row, 76].Text);//POP_2018_R
                    if (!workSheet.Cells[row, 392].Text.Equals("--")) s267 = Double.Parse(workSheet.Cells[row, 392].Text);//PRO_COM_110
                    if (!workSheet.Cells[row, 14].Text.Equals("--")) s268 = Double.Parse(workSheet.Cells[row, 14].Text);
                    if (!workSheet.Cells[row, 15].Text.Equals("--")) s269 = Double.Parse(workSheet.Cells[row, 15].Text);
                    if (!workSheet.Cells[row, 16].Text.Equals("--")) s270 = Double.Parse(workSheet.Cells[row, 16].Text);
                    if (!workSheet.Cells[row, 17].Text.Equals("--")) s271 = Double.Parse(workSheet.Cells[row, 17].Text);
                    if (!workSheet.Cells[row, 18].Text.Equals("--")) s272 = Double.Parse(workSheet.Cells[row, 18].Text);
                    if (!workSheet.Cells[row, 19].Text.Equals("--")) s273 = Double.Parse(workSheet.Cells[row, 19].Text);
                    if (!workSheet.Cells[row, 55].Text.Equals("--")) s274 = Double.Parse(workSheet.Cells[row, 55].Text);
                    if (!workSheet.Cells[row, 56].Text.Equals("--")) s275 = Double.Parse(workSheet.Cells[row, 56].Text);
                    if (!workSheet.Cells[row, 57].Text.Equals("--")) s276 = Double.Parse(workSheet.Cells[row, 57].Text);
                    if (!workSheet.Cells[row, 40].Text.Equals("--")) s277 = Double.Parse(workSheet.Cells[row, 40].Text);
                    if (!workSheet.Cells[row, 41].Text.Equals("--")) s278 = Double.Parse(workSheet.Cells[row, 41].Text);
                    if (!workSheet.Cells[row, 42].Text.Equals("--")) s279 = Double.Parse(workSheet.Cells[row, 42].Text);
                    if (!alreadyExists(comparator))
                    {
                        insertDB(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21, s22, s23, s24, s25, s26, s27, s28, s29, s30, s31, s32, s33, s34, s35, s36, s37, s38, s39, s40, s41, s42, s43, s44, s45, s46, s47, s48, s49, s50, s51, s52, s53, s54, s55, s56, s57, s58, s59, s60, s61, s62, s63, s64, s65, s66, s67, s68, s69, s70, s71, s72, s73, s74, s75, s76, s77, s78, s79, s80, s81, s82, s83, s84, s85, s86, s87, s88, s89, s90, s91, s92, s93, s94, s95, s96, s97, s98, s99, s100, s101, s102, s103, s104, s105, s106, s107, s108, s109, s110, s111, s112, s113, s114, s115, s116, s117, s118, s119, s120, s121, s122, s123, s124, s125, s126, s127, s128, s129, s130, s131, s132, s133, s134, s135, s136, s137, s138, s139, s140, s141, s142, s143, s144, s145, s146, s147, s148, s149, s150, s151, s152, s153, s154, s155, s156, s157, s158, s159, s160, s161, s162, s163, s164, s165, s166, s167, s168, s169, s170, s171, s172, s173, s174, s175, s176, s177, s178, s179, s180, s181, s182, s183, s184, s185, s186, s187, s188, s189, s190, s191, s192, s193, s194, s195, s196, s197, s198, s199, s200, s201, s202, s203, s204, s205, s206, s207, s208, s209, s210, s211, s212, s213, s214, s215, s216, s217, s218, s219, s220, s221, s222, s223, s224, s225, s226, s227, s228, s229, s230, s231, s232, s233, s234, s235, s236, s237, s238, s239, s240, s241, s242, s243, s244, s245, s246, s247, s248, s249, s250, s251, s252, s253, s254, s255, s256, s257, s258, s259, s260, s261, s262, s263, s264, s265, s266, s267, s268, s269, s270, s271, s272, s273, s274, s275, s276, s277, s278, s279);
                        Console.WriteLine("Record inserito riga " + row);
                        n_inseriti++;
                    }
                }
            }
            return n_inseriti;
        }
    }
}
