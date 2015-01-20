using CRM.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace HAT_API_PRD
{
    class PrdModel
    {
        //static
        private OrganizationServiceContext xrm = EnvironmentSetting.Xrm;
        private IOrganizationService service = EnvironmentSetting.Service;

        //special
        private int uom;
        private int new_sunit;

        //lookup Guid
        private static String[] lookupNameArray = { "gprdno", "prcno", "prlno"
                                                      , "prmno", "prhno", "prpno" };
        private int[] lookupIntArray = new int[lookupNameArray.Length];

        private Guid uomscheduleid;
        private Guid uomid;

        //lookup

        //special
        private int productnumber;
        private int name;

        //int
        private static String[] intNameArray = { "pakn", "pakhp" };
        private int[] intIntArray = new int[intNameArray.Length];

        //decimal
        private static String[] decimalNameArray = { "eprice", "hprice", "lprice"
                                                       ,"gpdprice","price" };
        private int[] decimalIntArray = new int[decimalNameArray.Length];

        //string
        private static String[] stringNameArray = {"pakspc","component","prdena"
                                                      ,"sayn","insuitem","p123"
                                                      ,"prdcna","pmark","ingride"
                                                      ,"ean13","invitna","prhlicno"
                                                      ,"ctrlcla","hpitem","prdmcc"};
        private int[] stringIntArray = new int[stringNameArray.Length];

        public PrdModel(SqlDataReader reader)
        {
            try
            {
                //lookup
                for (int i = 0, size = lookupNameArray.Length; i < size; i++)
                {
                    lookupIntArray[i] = reader.GetOrdinal(lookupNameArray[i]);
                }

                uom = reader.GetOrdinal("unit");
                new_sunit = reader.GetOrdinal("sunit");

                uomscheduleid = Lookup.RetrieveEntityGuid("uomschedule", EnvironmentSetting.UomScheduleName.ToString(), "name");
                uomid = Lookup.RetrieveEntityGuid("uom", EnvironmentSetting.Uom.ToString(), "name");

                //special
                productnumber = reader.GetOrdinal("prdno");
                name = reader.GetOrdinal("prdna");

                //int
                for (int i = 0, size = intNameArray.Length; i < size; i++)
                {
                    intIntArray[i] = reader.GetOrdinal(intNameArray[i]);
                }

                //decimal
                for (int i = 0, size = decimalNameArray.Length; i < size; i++)
                {
                    decimalIntArray[i] = reader.GetOrdinal(decimalNameArray[i]);
                }

                //string
                for (int i = 0, size = stringNameArray.Length; i < size; i++)
                {
                    stringIntArray[i] = reader.GetOrdinal(stringNameArray[i]);
                }
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg += "搜尋欄位失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DB;
            }

        }
        public Guid IsProductExist(SqlDataReader reader)
        {
            try
            {
                return Lookup.RetrieveEntityGuid("product", reader.GetString(productnumber).Trim(), "productnumber");
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg += "檢查資料失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNCDETAIL;
                return Guid.Empty;
            }
        }
        public TransactionStatus CreateProductForCRM(SqlDataReader reader)
        {
            return AddAttributeForRecord(reader, Guid.Empty);
        }
        public TransactionStatus UpdateProductForCRM(SqlDataReader reader, Guid productId)
        {
            return AddAttributeForRecord(reader, productId);
        }
        private TransactionStatus AddAttributeForRecord(SqlDataReader reader, Guid entityId)
        {
            try
            {
                Entity entity = new Entity("product");

                //special
                entity["productnumber"] = reader.GetString(productnumber).Trim();
                entity["name"] = reader.GetString(name).Trim();

                //int
                for (int i = 0, size = intNameArray.Length; i < size; i++)
                {
                    entity["new_" + intNameArray[i]] = reader.GetInt32(intIntArray[i]);
                }
                //decimal
                for (int i = 0, size = decimalNameArray.Length; i < size; i++)
                {
                    entity["new_" + decimalNameArray[i]] = new Money(reader.GetDecimal(decimalIntArray[i]));
                }
                //string
                for (int i = 0, size = stringNameArray.Length; i < size; i++)
                {
                    entity["new_" + stringNameArray[i]] = reader.GetString(stringIntArray[i]).Trim();
                }

                //lookup
                String recordStr;
                Guid recordGuid;
                //這個是產品內建必填欄位
                entity["defaultuomscheduleid"] = new EntityReference("uomschedule", uomscheduleid);

                /// 這個是必填 所以空值有給預設ID
                /// 
                /// CRM欄位名稱     銷售單位    defaultuomid
                /// CRM關聯實體     單位        uom
                /// CRM關聯欄位     名曾        name
                /// ERP欄位名稱                 unit
                /// 
                recordStr = reader.GetString(uom).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["defaultuomid"] = new EntityReference("uom", uomid);
                else
                {
                    recordGuid = Lookup.RetrieveEntityGuid("uom", recordStr, "name");
                    if (recordGuid == Guid.Empty)
                    {
                        EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                        EnvironmentSetting.ErrorMsg += "\tCRM實體 : uom\n";
                        EnvironmentSetting.ErrorMsg += "\tERP欄位 : unit\n";
                        //EnvironmentSetting.ErrorMsg += "\tERP資料 : " + recordStr + "\n";
                        Console.WriteLine(EnvironmentSetting.ErrorMsg);
                        return TransactionStatus.Fail;
                    }
                    entity["defaultuomid"] = new EntityReference("uom", recordGuid);
                }

                /// CRM欄位名稱     包裝小單位    defaultuomid
                /// CRM關聯實體     單位          uom
                /// CRM關聯欄位     名曾          name
                /// ERP欄位名稱                   sunit
                /// 
                recordStr = reader.GetString(new_sunit).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["new_sunit"] = null;
                else
                {
                    recordGuid = Lookup.RetrieveEntityGuid("uom", recordStr, "name");
                    if (recordGuid == Guid.Empty)
                    {
                        EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                        EnvironmentSetting.ErrorMsg += "\tCRM實體 : uom\n";
                        EnvironmentSetting.ErrorMsg += "\tERP欄位 : sunit\n";
                        //EnvironmentSetting.ErrorMsg += "\tERP資料 : " + recordStr + "\n";
                        Console.WriteLine(EnvironmentSetting.ErrorMsg);
                        return TransactionStatus.Fail;
                    }
                    entity["new_sunit"] = new EntityReference("uom", recordGuid);
                }

                /// CRM欄位名稱     "new_" + lookupNameArray[i]
                /// CRM關聯實體     "new_" + lookupNameArray[i]
                /// CRM關聯欄位     "new_" + lookupNameArray[i]
                /// ERP欄位名稱     lookupNameArray[i]
                /// 
                for (int i = 0, size = lookupNameArray.Length; i < size; i++)
                {
                    recordStr = reader.GetString(lookupIntArray[i]).Trim();
                    if (recordStr == "" || recordStr == null)
                        entity["new_" + lookupNameArray[i]] = null;
                    else
                    {
                        //prcno 是 CRM 和 ERP 欄位名稱不一樣 比較特別
                        if (lookupNameArray[i] == "prcno")
                        {
                            recordGuid = Lookup.RetrieveEntityGuid("new_prcn", recordStr, "new_prcn");
                            if (recordGuid == Guid.Empty)
                            {
                                EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                                EnvironmentSetting.ErrorMsg += "\tCRM實體 : new_prcn\n";
                                EnvironmentSetting.ErrorMsg += "\tERP欄位 : prcno\n";
                                //EnvironmentSetting.ErrorMsg += "\tERP資料 : " + recordStr + "\n";
                                Console.WriteLine(EnvironmentSetting.ErrorMsg);
                                return TransactionStatus.Fail;
                            }
                            entity["new_prcno"] = new EntityReference("new_prcn", recordGuid);
                        }
                        else
                        {
                            recordGuid = Lookup.RetrieveEntityGuid("new_" + lookupNameArray[i]
                                                                , recordStr
                                                                , "new_" + lookupNameArray[i]);
                            if (recordGuid == Guid.Empty)
                            {
                                EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                                EnvironmentSetting.ErrorMsg += "\tCRM實體 : new_" + lookupNameArray[i] + "\n";
                                EnvironmentSetting.ErrorMsg += "\tERP欄位 : " + lookupNameArray[i] + "\n";
                                //EnvironmentSetting.ErrorMsg += "\tERP資料 : " + recordStr + "\n";
                                Console.WriteLine(EnvironmentSetting.ErrorMsg);
                                return TransactionStatus.Fail;
                            }
                            entity["new_" + lookupNameArray[i]] = new EntityReference("new_" + lookupNameArray[i], recordGuid);
                        }
                    }
                }

                try
                {
                    if (entityId == Guid.Empty)
                        service.Create(entity);
                    else
                    {
                        entity["productid"] = entityId;
                        service.Update(entity);
                    }
                    return TransactionStatus.Success;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("productnumber : " + reader.GetString(productnumber).Trim());
                    Console.WriteLine(ex.Message);
                    EnvironmentSetting.ErrorMsg = ex.Message;
                    return TransactionStatus.Fail;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("欄位讀取錯誤");
                Console.WriteLine(ex.Message);
                EnvironmentSetting.ErrorMsg = "欄位讀取錯誤\n" + ex.Message;
                return TransactionStatus.Fail;
            }
        }
    }
}
