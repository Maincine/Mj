import pymssql
import time
import xlrd
import pymysql
def read_excel():
    # 打开excel文件读取数据
    data = xlrd.open_workbook("D:\ksf.xlsx")
    # 获取book中所有工作表的名称
    print("获取book中所有工作表的名称:")
    print(data.sheet_names())
    # ['Sheet1']

    # 根据工作表的名称获取工作表的内容
    table = data.sheet_by_name('Sheet1')

    # 根据工作表的索引获取工作表的内容
    # table = data.sheet_by_name(0)

    # 打印工作表的名称、行数和列数
    print("打印工作表的名称、行数和列数:")
    print(table.name, table.nrows, table.ncols)
    nrows=table.nrows
    list2=[]
    for i in range(nrows):
        arr=table.row_values(i)
        arr=list(map(int,arr))

        list2.append(arr)

    return list2

def bpmbase(listinfo):
    conn = pymssql.connect(server="10.0.101.19", user="FM", password="qaz123.1", database="DingXinBiz", charset='utf8')
    print(conn)
    cursor = conn.cursor()
    info = []
    for mastercode in listinfo:

        sql="SELECT STORENAME	,STORECODE,STORECOUNTYID,STOREADDRESS,STOREPHONE1, LONGITUDE,LATITUDE ,AREACOMPANYID,SUBCOMPANYID,BUSINESSCENTERID,SUBCOMPANY FROM MNG_DICOS_MF_MASTERKONGINFO where storecode ={}".format (mastercode[1])
        cursor.execute(sql)

        result = cursor.fetchall()

        result = list(result[0])
        # result.append(mastercode[0])
        result.insert(0, mastercode[0])
        info.append(result)


    conn.commit()
    cursor.close()
    conn.close()
    print("需要插入的餐厅信息",info)
    return info

def  insert_store(info):
    conn = pymysql.connect(host="172.16.3.29"
                           , user="root",
                           password="dicos8888",
                           database="ddos_misc",
                           charset="utf8")
    cursor = conn.cursor()


    for masterinfo in data:
        sqlbaseinfo="INSERT INTO `ddos_misc`.`master_store_t_base_info` (`GUID`, `STORE_CODE`, `FIN_STORE_CODE`, `CXJ_STORE_CODE`, `STORE_NAME_CN`, `STORE_NAME_EN`, `ALLIANCES`, `OPERATE_STATE`, `JOIN_OWNER`, `EMP_RECORD`, `RECORD_PLAN`, `RECORD_DATE_FROM`, `RECORD_DATE_TO`, `STORE_TYPE`, `IS_CHANGE`, `CHANGE_TYPE`, `CHANGE_DATE_FROM`, `CHANGE_DATE_TO`, `REGIONALISM_CODE`, `STORE_ADDRESS`, `STORE_PHONE`, `STORE_MAIL`, `STORE_POSTAL`, `SALE_CHANNEL`, `COMPETE`, `CITY_TYPE`, `BD_TYPE`, `STORE_FLOOR`, `STORE_AREA`, `STORE_SEATS`, `KOISK_NUM`, `LONGITUDE`, `LATITUDE`, `PARENT_STORE_CODE`, `POS_COUNT`, `BANDWIDTH`, `BANDWIDTHACCOUNT`, `BANDWIDTHPASSWORD`, `SELF_MACHINE_COUNT`, `IS_COMPETE`, `IS_EQU_LEASE`, `IS_MUSLIM_STORE`, `IS_PLAYGROUND`, `IS_EXTERNALLEASE`, `IS_SELLCARD`, `IS_24HOUR`, `IS_CALL_ORDER`, `IS_ELECTRONICLIGHT`, `IS_WIFI`, `IS_ISSPOT`, `ISOFFSTANDARDSTORE`, `IS_E_INVOICE`, `IS_EQU_COST`, `IS_PLIOCY_SUPPORT`, `TAXPAYER_TYPE`, `FINDEVICE_TAX`, `FINBRANDTAX`, `LEGAL_PERSON`, `INVOICE_COMPANY_NAME`, `INVOICE_TAXNUMBER`, `INVOICE_ADDRESS`, `INVOICE_PHONE`, `INVOICE_OPENING_BANK`, `INVOICE_ACCOUNT`, `SHIPPING_COMPANY`, `MATERIEL_SERVICE_AMOUNT`, `ELECTRI_CITY_PRICE`, `WATER_PRICE`, `JOIN_AMOUNT`, `ASSURE_AMOUNT`, `SIGNATURE_AMOUNT`, `IS_PLIOCY_SUPPORT_JM`, `PLIOCY_SUPPORT_JM_DATE_FROM`, `PLIOCY_SUPPORT_JM_DATE_TO`, `JOIN_CONTRACT_DATE_FROM`, `JOIN_CONTRACT_DATE_TO`, `PLIOCY_SUPPORT_QT_OTHER`, `HOUSINGRENTAL_DATE_FROM`, `HOUSINGRENTAL_DATE_TO`, `STORE_RENTAL`, `STORE_RENTAL_TYPE`, `STORE_RENTAL_CUT`, `STORE_RENTAL_ESTATEMGR`, `STORE_RENTAL_HOUSE`, `STORE_RENTAL_OUT`, `STORE_RENTAL_REMARK`, `CALL_SCREEN_TYPE`, `CP_URL`, `STATE_FLAG`, `BRAND_CODE`, `CREATE_ID`, `CREATE_TIME`, `UPDATE_ID`, `UPDATE_TIME`, `DEL_FLAG`, `SOURCE`, `IS_ONLINE`, `appKey`, `appSecret`, `linkAddress`) VALUES (uuid(), '{}', '{}', NULL, '{}', NULL, '1', '5', '业主', '0', NULL, NULL, NULL, '2', '0', NULL, NULL, NULL, '{}', '{}', '{}', NULL, NULL, NULL, NULL, '3', '1', '1', '400.00', '108', '1', '{}', '{}', NULL, '4台', '8', 'dicosghddy', '5229099', '1', '0', NULL, '0', '1', '1', '1', NULL, NULL, '0', '1', '0', '0', NULL, '0', NULL, NULL, NULL, '3', NULL, NULL, NULL, NULL, NULL, NULL, NULL, 'B230600', '1.1', '2.00', '3.50', '3', '9', '1', '0', NULL, NULL, NULL, NULL, NULL, '2014-09-01 00:00:00', '2019-08-31 00:00:00', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2', 'admin', '2017-11-09 11:38:54', 'admin', '2017-11-09 11:38:54', 'N', NULL, 'Y', NULL, NULL, 'http://weixin.qq.com/q/02g8GCZ6E9bWe100000075');".format(masterinfo[0],masterinfo[2],masterinfo[1],masterinfo[3],masterinfo[4],masterinfo[5],masterinfo[6],masterinfo[7])
        sqldelivery="INSERT INTO `ddos_misc`.`master_store_t_delivery` (`GUID`, `STORE_CODE`, `IS_DELIVERY_SPREAD`, `DELIVERY_TIME_START`, `DELIVERY_TIME_END`, `DELIVERY_PRICE`, `DELIVERY_AREA`, `DELIVERY_EACH_TIME`, `DELIVERY_STANDARD`, `DELIVERY_FREE_PRICE`, `DELIVERY_CHARGE_STANDARD`, `DELIVERY_TYPE`, `DELIVERY_BUSINESS`, `IS_DELIVERY_DISTRICT`, `BRAND_CODE`, `CREATE_ID`, `CREATE_TIME`, `UPDATE_ID`, `UPDATE_TIME`, `DEL_FLAG`) VALUES (uuid(), '{}', '1', '08:00:00', '19:00:00', '20', '1', '30', '1', '30', '30', '1,2,3,4', '1,2,3,0', '1', '2', 'admin', '2018-03-26 10:15:35', 'admin', '2018-03-26 10:15:35', 'N');".format(masterinfo[0])
        sqlbussiness="INSERT INTO `ddos_misc`.`master_store_t_business_time` (`GUID`, `STORE_CODE`, `START_BUSINESS_DATE`, `END_BUSINESS_DATE`, `IE_TYPE_CODE`, `IE_START_DATE`, `IE_END_DATE`, `OPERATING_MONTH`, `OPERATING_WEEK`, `OPERATING_TIME_START`, `OPERATING_TIME_END`, `IS_BREAKFAST`, `BREAKFAST_TIME_START`, `BREAKFAST_TIME_END`, `BRAND_CODE`, `CREATE_ID`, `CREATE_TIME`, `UPDATE_ID`, `UPDATE_TIME`, `DEL_FLAG`) VALUES (uuid(), '{}', '2012-05-24 00:00:00', NULL, '1', NULL, NULL, '1,2,3,4,5,6,7,8,10,11,12', '1,2,3,4,5,6,7', '07:00:00', '22:30:00', '1', '07:00:00', '10:00:00', '2', 'admin', '2017-11-09 11:39:05', 'admin', '2017-11-09 12:23:57', 'N');".format(masterinfo[0])
        sqlmarket="INSERT INTO `ddos_misc`.`master_store_t_org_tree` (`GUID`, `STORE_CODE`, `AREA_COMPANY`, `SUB_COMPANY`, `BUSINESS_CENTER`, `AREA_SUPERVISOR_AREA`, `AREA_SUPERVISOR`, `AREA_SUPERVISOR_PHONE`, `RGM`, `RGM_PHONE`, `STATE_FLAG`, `BRAND_CODE`, `CREATE_ID`, `CREATE_TIME`, `UPDATE_ID`, `UPDATE_TIME`, `DEL_FLAG`) VALUES (uuid(), '{}', '{}', '{}', '{}', 'B003',  SUBSTRING('{}',1,6), '13704507427', '{}', NULL, NULL, '2', 'admin', '2018-03-26 10:15:31', 'admin', '2018-03-26 10:15:31', 'N');".format(masterinfo[0],masterinfo[8],masterinfo[9],masterinfo[10],masterinfo[9],masterinfo[0])
        sqluser="INSERT INTO `ddos_misc`.`master_fund_t_system_user` (`guid`, `user_id`, `user_name`, `password`, `email`, `mobile`, `status`, `create_time`, `create_id`, `update_time`, `update_id`, `del_flag`, `area_code`, `store_code`, `itsm_create_user_flag`, `user_type`, `out_system_user_id`, `uplevel_user_id`, `brand_code`) VALUES (uuid(), '{}', '{}', '8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918', 'fei_xiang@hoperun.com', NULL, '1', '2017-11-09 12:49:24.000000', 'admin', '2017-11-09 12:49:24.000000', 'admin', 'N', NULL, NULL, '1', '2', NULL, NULL, '2');".format(masterinfo[0],masterinfo[0])
        sqlrole="INSERT INTO `ksf_mobile`.`master_fund_t_system_user_role` (`guid`, `user_id`, `role_id`, `create_time`, `create_id`, `update_time`, `update_id`) VALUES (uuid(), '{}', 'R00012', '2019-07-08 09:24:25.831844', 'admin', '2019-07-04 13:34:56.000000', 'admin');".format(masterinfo[0])
        print(sqlbaseinfo)
        print(sqldelivery)
        print(sqlbussiness)
        print(sqlmarket)
        print(sqluser)
        print(sqlrole)
        try:
            cursor.execute(sqlbaseinfo)
            cursor.execute(sqlbussiness)
            cursor.execute(sqldelivery)
            cursor.execute(sqlmarket)
            cursor.execute(sqluser)
            cursor.execute(sqlrole)
        except Exception as e:
            conn.rollback()
            print('事务处理失败', e)
        else:
            conn.commit()  # 事务提交cursor.close()
            print("suess instert")
    sqlmobile = "INSERT INTO ksf_mobile.master_fund_t_system_user (guid,user_id,user_name,PASSWORD,email,mobile,STATUS,create_time,create_id,update_time,update_id,del_flag,store_code,user_type) SELECT a.guid,a.STORE_CODE,a.STORE_NAME_CN,'5456ab660970ea7468ffc61bbf5566fab9ba4e7e98890f8465f62214ea2c8674',a.STORE_MAIL,a.STORE_PHONE,'1',SYSDATE(),'mj',SYSDATE(),'mj',a.DEL_FLAG,a.STORE_CODE,'2' FROM ddos_misc.master_store_t_base_info a LEFT JOIN ksf_mobile.master_fund_t_system_user b ON a.STORE_CODE = b.user_id AND b.del_flag = 'N' WHERE b.user_id IS NULL AND a.DEL_FLAG = 'N' AND a.brand_code = 2;"
    print(sqlmobile)
    try:
        cursor.execute(sqlmobile)
    except Exception as e:
        conn.rollback()
        print('事务处理失败', e)
    else:
        conn.commit()  # 事务提交cursor.close()
        print("suess instert")
    cursor.close()
    conn.close

def instert_ksf_ios(info):
    conn =pymysql.connect(host="172.16.3.29"
                           , user="root",
                           password="dicos8888",
                           database="dicos_ios",
                           charset="utf8")
    cursor = conn.cursor()

    for masterinfo in data:
        sql="INSERT INTO `ksf_ios`.`store` (`ID`, `StoreCode`, `FinStoreCode`, `CXJStoreCode`, `StoreName`, `StoreAddressPy`, `Address`, `DistrictCode`, `ZipCode`, `email`, `IP`, " \
            "`Contact`, `Phone`, `Property`, `Status`, `Delivery`, `Traffic`, `PosVersion`, `IsUsed`, `storeType`, `enableEmail`, `MenuVersion`, `RapVersion`, `PosMenuVersion`, `PosRapVersion`," \
            " `LastMenuSerial`, `Port`, `DeliveryTime`, `BusinessDate`, `ForceSendByEmail`, `ShutReason`, `Merchantno`, `starttime`, `endtime`, `OpenType`, `opendate`, `closedate`, `openOnlinePay`, `abbr`, `ShutMapTime`, `phone2`, `RGM`, `AM`, `DM`, `isCall`, `coordinate_x`, `coordinate_y`, `repairOrderEmail`, `isHui`, `isEInvoice`, `isInvoiceTitle`, `outofInvoice`, `storeTypes`, `channelId`, `appKey`, `appSecret`, `posType`, `isEnclosing`, `deliveryPrice`, `businessLicence`, `businessCertificate`, `storeShowName`, `MarketCode`, `MarketName`, `brandCode`, `boxPrice`, `invoiceFirm`, `encryptKey`, `taxNo`, `eInvoiceUrl`, `invoiceClass`, `invoiceNo`) " \
            "VALUES ('{}', '{}', '{}', NULL, '{}', NULL, '{}', '{}', NULL, NULL, NULL, NULL, '{}', NULL, '1', NULL, NULL, NULL, NULL, '2', NULL, NULL, NULL, NULL, NULL, CONCAT('DDOS_MENU_', '{}')," \
            " NULL, NULL, NULL, NULL, NULL, '1', '07:30:00', '21:00:00', '1111', NULL, NULL, '1', NULL, NULL, NULL, '{}', SUBSTRING('{}',1,6), NULL, NULL, '{}', '{}', NULL, NULL, '1', '0', NULL, '00000001', '11'," \
            " NULL, Null , '2', '1', '23', NULL, NULL, '{}', '{}'," \
            "'{}',  '1', '4', NULL, NULL, NULL, NULL, NULL, NULL);".format(masterinfo[0],masterinfo[0],masterinfo[2],masterinfo[1],masterinfo[4],masterinfo[3],masterinfo[5],masterinfo[0],masterinfo[0],masterinfo[9],masterinfo[6],masterinfo[7],masterinfo[1],masterinfo[9],masterinfo[11])
        sql1  ="INSERT INTO `ksf_ios`.`t_storechannel` ( `channel`, `storeCode`, `startTime`, `endTime`, `deliveryType`, `status`, `shutReason`, `shutMapTime`, `consignerId`, `consignerShowName`, `serviceFeeRate`) VALUES ( '1', '{}', '20:30', '23:00', '2', '1', '', NULL, '', '', '1');"
        sql2  ="INSERT INTO `ksf_ios`.`t_storechannel` ( `channel`, `storeCode`, `startTime`, `endTime`, `deliveryType`, `status`, `shutReason`, `shutMapTime`, `consignerId`, `consignerShowName`, `serviceFeeRate`) VALUES ( '3', '210028', '08:21', '23:00', '1', '1', '', NULL, '', '', NULL);"
        sql3  = "INSERT INTO `ksf_ios`.`t_storechannel` ( `channel`, `storeCode`, `startTime`, `endTime`, `deliveryType`, `status`, `shutReason`, `shutMapTime`, `consignerId`, `consignerShowName`, `serviceFeeRate`) VALUES ( '4', '210028', '08:30', '23:10', '1', '1', '', NULL, '', '', NULL);"
        sql4 =  "INSERT INTO `ksf_ios`.`t_storechannel` ( `channel`, `storeCode`, `startTime`, `endTime`, `deliveryType`, `status`, `shutReason`, `shutMapTime`, `consignerId`, `consignerShowName`, `serviceFeeRate`) VALUES ( '5', '210028', '06:37', '23:10', '1', '1', '', NULL, '', '', NULL);"
        sql5=" INSERT INTO `ksf_ios`.`t_storechannel` ( `channel`, `storeCode`, `startTime`, `endTime`, `deliveryType`, `status`, `shutReason`, `shutMapTime`, `consignerId`, `consignerShowName`, `serviceFeeRate`) VALUES ( '6', '210028', '06:37', '23:10', '2', '2', '', NULL, '', '', NULL);"











if __name__ == "__main__":
    return_list = read_excel()
    print(return_list)
    data = bpmbase(return_list)

    insert_store(data)
    #instert_ksf_ios(data)
