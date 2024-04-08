using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Patholab_DAL_V1;
using System.Xml;
using LSEXT;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Patholab_Common;
using LSSERVICEPROVIDERLib;

namespace AliquotXmlMaker
{
    [ComVisible(true)]
    [ProgId("AliquotXmlMaker.AliquotXmlMakerCls")]
    public class AliquotXmlMakerCls : IWorkflowExtension
    {
        INautilusServiceProvider sp;
        DataLayer dal;

        public void Execute(ref LSExtensionParameters Parameters)
        {
            //Debugger.Launch();

            string tableName = Parameters["TABLE_NAME"];

            sp = Parameters["SERVICE_PROVIDER"];
            var rs = Parameters["RECORDS"];


            int aliquotID = (int)rs["ALIQUOT_ID"].Value;
            //string NAME = rs["NAME"].Value.ToString();
            rs.MoveLast();
            //string tableID = rs.Fields["SDG_ID"].Value.ToString();
            //string workstationId = Parameters["WORKSTATION_ID"].ToString();
            //long sdgId = long.Parse(tableID);
            var ntlCon = Utils.GetNtlsCon(sp);
            Utils.CreateConstring(ntlCon);
            /////////////////////////////           
            dal = new DataLayer();
            dal.Connect(ntlCon);


            ALIQUOT aliquot = dal.FindBy<ALIQUOT>(x => x.ALIQUOT_ID == aliquotID).FirstOrDefault();

            PHRASE_HEADER phraseHeader = dal.GetPhraseByName("System Parameters");
            //PHRASE_HEADER phraseHeader = dal.FindBy<PHRASE_HEADER>(x => x.PHRASE_ID == 282).FirstOrDefault();
            //PHRASE_ENTRY phraseEntry = phraseHeader.PHRASE_ENTRY.Where(x => x.PHRASE_NAME == "Create XML").FirstOrDefault();
            string XmlDir;
            //Get xml destination path
            phraseHeader.PhraseEntriesDictonary.TryGetValue("Create XML", out XmlDir);

            string fullPath = XmlDir + aliquotID.ToString() + ".xml";

            //long aliquotID = aliquot.ALIQUOT_ID;
            createAliquotXML(aliquotID, fullPath);
        }

        public void createAliquotXML(long ID, string fullPath)
        {
            ALIQUOT aliquot = dal.FindBy<ALIQUOT>(x => x.ALIQUOT_ID == ID).FirstOrDefault();
            CLIENT_USER client = aliquot.SAMPLE.SDG.SDG_USER.CLIENT.CLIENT_USER;
            SUPPLIER_USER supplier;
            if(aliquot.SAMPLE.SDG.SDG_USER.REFERRING_PHYSIC != null){
                supplier = aliquot.SAMPLE.SDG.SDG_USER.REFERRING_PHYSIC.SUPPLIER_USER;
            } 
            else
            {
                supplier=null;
            }

            SAMPLE sample = aliquot.SAMPLE;
            U_ORDER_USER customer = aliquot.SAMPLE.SDG.SDG_USER.U_ORDER.U_ORDER_USER;

             



            using (XmlWriter writer = XmlWriter.Create(fullPath))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Demographic_Data");
                writer.WriteElementString("Id_type", client.U_PASSPORT);
                writer.WriteElementString("Id", client.CLIENT_ID.ToString());
                writer.WriteElementString("First_Name", client.U_FIRST_NAME);
                writer.WriteElementString("Last_Name", client.U_LAST_NAME);
                writer.WriteElementString("Father_Name", client.U_FATHER_NAME != null ? client.U_FATHER_NAME : "");
                writer.WriteElementString("Sex", client.U_GENDER);
                writer.WriteElementString("BDate", client.U_DATE_OF_BIRTH.ToString());
                writer.WriteElementString("Sick_Fund","");// קוד קופת חולים
                writer.WriteElementString("Sick_Fund_Desc","");// שם קופת חולים
                writer.WriteElementString("State","");// קוד ארץ מוצא
                writer.WriteElementString("State_Name","");// שם ארץ מוצא
                writer.WriteElementString("City_NO","");// קוד עיר
                writer.WriteElementString("City_name","");// שם עיר
                writer.WriteElementString("Street","");// שם רחוב
                writer.WriteElementString("Zip","");// מיקוד
                writer.WriteElementString("HTel","");// מספר טלפון בבית
                writer.WriteElementString("WTel", client.U_PHONE);
                writer.WriteElementString("Fax","");// מספר פקס
                writer.WriteElementString("Email","");// חשבון מייל
                writer.WriteElementString("En_First_Name","");// שם פרטי באנגלית
                writer.WriteElementString("En_Last_Name","");// שם משפחה באנגלית
                writer.WriteElementString("Immigration_Country","");// ארץ מוצא
                writer.WriteElementString("Immigration_Date","");// תאריך עלייה
                writer.WriteElementString("Confidential_Patient","");// מטופל חסוי [T][F]
                writer.WriteElementString("Macabi_No","");// מספר מכבי
                writer.WriteElementString("Polisa_Num","");// מספר פוליסה
                writer.WriteElementString("Death_date","");// תאריך פטירה

                writer.WriteStartElement("Patient_Visit");
                writer.WriteElementString("Type_Of_Order", sample.SDG.NAME[0].ToString());
                writer.WriteElementString("Order_No", aliquot.SAMPLE.SDG.SDG_USER.U_ORDER.U_ORDER_ID.ToString());
                writer.WriteElementString("Sample_NO", sample.SAMPLE_ID.ToString());
                writer.WriteElementString("Aliquot_NO", aliquot.NAME);
                writer.WriteElementString("Block_No", aliquot.NAME);
                writer.WriteElementString("Slide_No", aliquot.NAME);
                writer.WriteElementString("SDG_NO", aliquot.SAMPLE.SDG.SDG_ID.ToString());
                writer.WriteElementString("Site_Code","");
                writer.WriteElementString("Creation_File_Date", aliquot.SAMPLE.SDG.CREATED_ON.ToString());
                writer.WriteElementString("Pay_No", customer.U_CUSTOMER1.U_CUSTOMER_ID.ToString());
                writer.WriteElementString("Pay_No_Desc", customer.U_CUSTOMER1.NAME);
                writer.WriteElementString("Visit_Date","");// תאריך לקיחת הדגימה
                writer.WriteElementString("Refer_Doc_Licsense",supplier != null ? supplier.U_LICENSE_NBR : "");
                writer.WriteElementString("Refer_Doc_Id", supplier != null ? supplier.U_ID_NBR.ToString() : "");
                writer.WriteElementString("Refer_Doc_Name", supplier != null ? supplier.U_LAST_NAME : "");
                writer.WriteElementString("Lab_Acceptation_Date", aliquot.SAMPLE.SDG.RECEIVED_ON != null ? aliquot.SAMPLE.SDG.RECEIVED_ON.ToString() : "");
                writer.WriteElementString("Sample_Acceptation_Date", aliquot.SAMPLE.RECEIVED_ON != null ? aliquot.SAMPLE.RECEIVED_ON.ToString() : "");
                writer.WriteElementString("Lab_cassette_Print_Date","");// תאריך הדפסת קסטות
                writer.WriteElementString("Lab_Mcro_Date", "");// תאריך ביצוע עמדת מאקרו
                writer.WriteElementString("Lab_Tissue_processing_Date","");// תאריך ביצוע עמדת מאקרו
                writer.WriteElementString("Lab_Tissue_setting_Date","");// תאריך עמדת שיקוע
                writer.WriteElementString("Lab_slide_setting_Date", aliquot.CREATED_ON.ToString());
                writer.WriteElementString("Lab_slide_coloring_Date","");// תאריך עמדת צביעה
                writer.WriteElementString("Lab_slide_digital_Date","");// תאריך הכנסת סלייד לסורק
                writer.WriteElementString("Doc_License", aliquot.SAMPLE.SDG.SDG_USER.PATHOLOG.OPERATOR_USER.U_LICENSE_NBR);
                writer.WriteElementString("Doc_Id",aliquot.SAMPLE.SDG.SDG_USER.PATHOLOG.OPERATOR_USER.OPERATOR_ID.ToString());
                writer.WriteElementString("Doc_Name", aliquot.SAMPLE.SDG.SDG_USER.PATHOLOG.FULL_NAME);
                writer.WriteElementString("Doc_Title_Desc", aliquot.SAMPLE.SDG.SDG_USER.PATHOLOG.OPERATOR_USER.U_DEGREE);
                writer.WriteElementString("Secertery_Name","");// שם מזכירה מקבלת
                writer.WriteElementString("Lab_Tech_Name_Sample_Acceptation","");// שם טכנאי מקבל צנצנות
                writer.WriteElementString("Lab_Tech_Name_cassette_Print","");// שם טכנאי הדפסת קסטות
                writer.WriteElementString("Lab_Tech_Name_Mcro","");// שם טכנאי ביצוע מאקרו
                writer.WriteElementString("Lab_Name_Tech_Tissue_processing","");// שם טכנאי עיבוד רקמות
                writer.WriteElementString("Lab_Name_Tech_Tissue_setting","");// שם טכנאי עמדת שיקוע
                writer.WriteElementString("Lab_Name_Tech_slide_setting","");// שם טכנאי חיתוך סליידים
                writer.WriteElementString("DiagName","");// שם האבחנה
                writer.WriteElementString("DiagCode","");// קוד האבחנה
                writer.WriteElementString("DiagMod","");// סטטוס האבחנה
                writer.WriteElementString("DiagDetail",""); //פרטי האבחנה
                writer.WriteEndElement();

                writer.WriteStartElement("audit");
                writer.WriteElementString("last_updated_by_user_id","");// יוזר אחרון שביצע שינוי
                writer.WriteElementString("last_updated_Date","");// תאריך עדכון אחרון של מסר נוכחי
                writer.WriteElementString("patient_key", client.CLIENT_ID.ToString());
                writer.WriteElementString("order_key", aliquot.SAMPLE.SDG.SDG_USER.U_ORDER.U_ORDER_ID.ToString());
                writer.WriteElementString("study_key", aliquot.SAMPLE.SDG.SDG_ID.ToString());
                writer.WriteElementString("HitNum","");// מספר התחייבות
                writer.WriteEndElement();

                writer.WriteEndDocument();
                writer.Flush();
                writer.Close();
            }

        }

    }
}
