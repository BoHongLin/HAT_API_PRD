using CRM.Common;
using CRM.Model;
using NLog;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace HAT_API_PRD
{
    class PrdFlow
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        public void DoPrdFlow()
        {
            DataSyncModel dataSync = new DataSyncModel();
            //connection CRM
            EnvironmentSetting.LoadSetting();
            if (EnvironmentSetting.ErrorType == ErrorType.None)
            {
                //create dataSync
                dataSync.CreateDataSyncForCRM("產品");
                if (EnvironmentSetting.ErrorType == ErrorType.None)
                {
                    //connection DB
                    EnvironmentSetting.LoadDB();
                    if (EnvironmentSetting.ErrorType == ErrorType.None)
                    {
                        //get information from DB
                        EnvironmentSetting.GetList("select * from dbo.v_prd");
                        if (EnvironmentSetting.ErrorType == ErrorType.None)
                        {
                            //get reader
                            SqlDataReader reader = EnvironmentSetting.Reader;

                            //search field index
                            PrdModel prdmodel = new PrdModel(reader);
                            if (EnvironmentSetting.ErrorType == ErrorType.None)
                            {
                                Console.WriteLine("連線成功!!");
                                Console.WriteLine("開始執行...");

                                if (reader.HasRows)
                                {
                                    int success = 0;
                                    int fail = 0;
                                    int partially = 0;
                                    while (reader.Read())
                                    {
                                        //判斷CRM是否有資料
                                        Guid existProductId = prdmodel.IsProductExist(EnvironmentSetting.Reader);
                                        if (EnvironmentSetting.ErrorType == ErrorType.None)
                                        {
                                            TransactionStatus transactionStatus;
                                            TransactionType transactionType;
                                            if (existProductId == Guid.Empty)
                                            {
                                                //create product
                                                transactionType = TransactionType.Insert;
                                                transactionStatus = prdmodel.CreateProductForCRM(EnvironmentSetting.Reader);
                                            }
                                            else
                                            {
                                                //update product
                                                transactionType = TransactionType.Update;
                                                transactionStatus = prdmodel.UpdateProductForCRM(EnvironmentSetting.Reader, existProductId);
                                            }

                                            if (EnvironmentSetting.ErrorType == ErrorType.None)
                                            {
                                                //create datasyncdetail
                                                switch (transactionStatus)
                                                {
                                                    case TransactionStatus.Success:
                                                        success++;
                                                        break;
                                                    case TransactionStatus.Fail:
                                                        //新增、更新資料有錯誤 則新增一筆detail
                                                        dataSync.CreateDataSyncDetailForCRM(reader["prdno"].ToString().Trim(), reader["prdna"].ToString().Trim(), transactionType, transactionStatus);
                                                        fail++;
                                                        break;
                                                    default:
                                                        fail++;
                                                        break;
                                                }

                                                //新增detail錯誤 則結束
                                                if (EnvironmentSetting.ErrorType != ErrorType.None)
                                                {
                                                    _logger.Info(EnvironmentSetting.ErrorMsg);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    //更新DataSync 成功、失敗、完成時間
                                    dataSync.UpdateDataSyncForCRM(success, fail, partially);
                                    Console.WriteLine("成功 : " + success);
                                    Console.WriteLine("失敗 : " + fail);
                                }
                                else
                                {
                                    Console.WriteLine("沒有資料");
                                    EnvironmentSetting.ErrorMsg += "ERP沒有資料\n";
                                    EnvironmentSetting.ErrorType = ErrorType.DB;
                                }
                            }
                        }
                    }
                }
            }
            switch (EnvironmentSetting.ErrorType)
            {
                case ErrorType.None:
                    break;

                case ErrorType.INI:
                case ErrorType.CRM:
                case ErrorType.DATASYNC:
                    _logger.Info(EnvironmentSetting.ErrorMsg);
                    break;

                case ErrorType.DB:
                case ErrorType.DATASYNCDETAIL:
                    dataSync.UpdateDataSyncWithErrorForCRM(EnvironmentSetting.ErrorMsg);
                    break;

                default:
                    break;
            }
            Console.WriteLine("執行完畢...");
            //Console.ReadLine();
        }
    }
}
