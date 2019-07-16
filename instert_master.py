import pymysql
import xlrd
def read_excel():
    # 打开excel文件读取数据
    data = xlrd.open_workbook("D:\mj.xlsx")

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



def instertmaster(listinfo):

    conn = pymysql.connect(host="172.16.3.29"
                       , user="root",
                       password="dicos8888",
                       database="ddos_misc",
                       charset="utf8")
    cursor = conn.cursor()


    # listinfo= list
    for mastercode in listinfo:

        sqlbaseinfo="INSERT INTO `ddos_misc`.`master_store_t_base_info` (GUID,\
	`STORE_CODE`,\
	`FIN_STORE_CODE`,\
	`CXJ_STORE_CODE`,\
	`STORE_NAME_CN`,\
	`STORE_NAME_EN`,\
	`ALLIANCES`,\
	`OPERATE_STATE`,\
	`JOIN_OWNER`,\
	`EMP_RECORD`,\
	`RECORD_PLAN`,\
	`RECORD_DATE_FROM`,\
	`RECORD_DATE_TO`,\
	`STORE_TYPE`,\
	`IS_CHANGE`,\
	`CHANGE_TYPE`,\
	`CHANGE_DATE_FROM`,\
	`CHANGE_DATE_TO`,\
	`REGIONALISM_CODE`,\
	`STORE_ADDRESS`,\
	`STORE_PHONE`,\
	`STORE_MAIL`,\
	`STORE_POSTAL`,\
	`SALE_CHANNEL`,\
	`COMPETE`,\
	`CITY_TYPE`,\
	`BD_TYPE`,\
	`STORE_FLOOR`,\
	`STORE_AREA`,\
	`STORE_SEATS`,\
	`KOISK_NUM`,\
	`LONGITUDE`,\
	`LATITUDE`,\
	`PARENT_STORE_CODE`,\
	`POS_COUNT`,\
	`BANDWIDTH`,\
	`BANDWIDTHACCOUNT`,\
	`BANDWIDTHPASSWORD`,\
	`SELF_MACHINE_COUNT`,\
	`IS_COMPETE`,\
	`IS_EQU_LEASE`,\
	`IS_MUSLIM_STORE`,\
	`IS_PLAYGROUND`,\
	`IS_EXTERNALLEASE`,\
	`IS_SELLCARD`,\
	`IS_24HOUR`,\
	`IS_CALL_ORDER`,\
	`IS_ELECTRONICLIGHT`,\
	`IS_WIFI`,\
	`IS_ISSPOT`,\
	`ISOFFSTANDARDSTORE`,\
	`IS_E_INVOICE`,\
	`IS_EQU_COST`,\
	`IS_PLIOCY_SUPPORT`,\
	`TAXPAYER_TYPE`,\
	`FINDEVICE_TAX`,\
	`FINBRANDTAX`,\
	`LEGAL_PERSON`,\
	`INVOICE_COMPANY_NAME`,\
	`INVOICE_TAXNUMBER`,\
	`INVOICE_ADDRESS`,\
	`INVOICE_PHONE`,\
	`INVOICE_OPENING_BANK`,\
	`INVOICE_ACCOUNT`,\
	`SHIPPING_COMPANY`,\
	`MATERIEL_SERVICE_AMOUNT`,\
	`ELECTRI_CITY_PRICE`,\
	`WATER_PRICE`,\
	`JOIN_AMOUNT`,\
	`ASSURE_AMOUNT`,\
	`SIGNATURE_AMOUNT`,\
	`IS_PLIOCY_SUPPORT_JM`,\
	`PLIOCY_SUPPORT_JM_DATE_FROM`,\
	`PLIOCY_SUPPORT_JM_DATE_TO`,\
	`JOIN_CONTRACT_DATE_FROM`,\
	`JOIN_CONTRACT_DATE_TO`,\
	`PLIOCY_SUPPORT_QT_OTHER`,\
	`HOUSINGRENTAL_DATE_FROM`,\
	`HOUSINGRENTAL_DATE_TO`,\
	`STORE_RENTAL`,\
	`STORE_RENTAL_TYPE`,\
	`STORE_RENTAL_CUT`,\
	`STORE_RENTAL_ESTATEMGR`,\
	`STORE_RENTAL_HOUSE`,\
	`STORE_RENTAL_OUT`,\
	`STORE_RENTAL_REMARK`,\
	`CALL_SCREEN_TYPE`,\
	`CP_URL`,\
	`STATE_FLAG`,\
	`BRAND_CODE`,\
	`CREATE_ID`,\
	`CREATE_TIME`,\
	`UPDATE_ID`,\
	`UPDATE_TIME`,\
	`DEL_FLAG`,\
	`SOURCE`,\
	`IS_ONLINE`,\
	`appKey`,\
	`appSecret`,\
	`linkAddress`\
) SELECT\
	UUID(),\
	'{}',\
	a.storecode,\
	NULL,\
	a.STORENAME,\
	NULL,\
	'1',\
	'5',\
	'业主',\
	'0',\
	NULL,\
	NULL,\
	NULL,\
	'2',\
	'0',\
	NULL,\
	NULL,\
	NULL,\
	a.storecountyid,\
	a.storeaddress,\
	a.storephone,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	'3',\
	'1',\
	'1',\
	'400.00',\
	'108',\
	'1',\
	a.longitude,\
	a.latitude,\
	NULL,\
	'4台',\
	'8',\
	'dicosghddy',\
	'5229099',\
	'1',\
	'0',\
	NULL,\
	'0',\
	'1',\
	'1',\
	'1',\
	NULL,\
	NULL,\
	'0',\
	'1',\
	'0',\
	'0',\
	NULL,\
	'0',\
	NULL,\
	NULL,\
	NULL,\
	'3',\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	'B230600',\
	'1.1',\
	'2.00',\
	'3.50',\
	'3',\
	'9',\
	'1',\
	'0',\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	'2014-09-01 00:00:00',\
	'2019-08-31 00:00:00',\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	NULL,\
	'1',\
	'admin',\
	'2017-11-09 11:38:54',\
	'admin',\
	'2017-11-09 11:38:54',\
	'N',\
	NULL,\
	'Y',\
	NULL,\
	NULL,\
	'http://weixin.qq.com/q/02g8GCZ6E9bWe100000075'\
FROM\
	bpm_master as a \
 WHERE a.storecode = {};".format(mastercode[0],mastercode[1])
        sqlbussiness="INSERT INTO `ddos_misc`.`master_store_t_business_time` \
(GUID ,\
	`STORE_CODE`,\
	`START_BUSINESS_DATE`,\
	`END_BUSINESS_DATE`,\
	`IE_TYPE_CODE`,\
	`IE_START_DATE`,\
	`IE_END_DATE`,\
	`OPERATING_MONTH`,\
	`OPERATING_WEEK`,\
	`OPERATING_TIME_START`,\
	`OPERATING_TIME_END`,\
	`IS_BREAKFAST`,\
	`BREAKFAST_TIME_START`,\
	`BREAKFAST_TIME_END`,\
	`BRAND_CODE`,\
	`CREATE_ID`,\
	`CREATE_TIME`,\
	`UPDATE_ID`,\
	`UPDATE_TIME`,\
	`DEL_FLAG`)\
VALUES\
		(UUID(),\
		{},\
		'2003-09-20 00:00:00',\
		NULL,\
		'1',\
		NULL,\
		NULL,\
		'1,2,3,4,5,6,7,8,9,10,11,12',\
		'1,2,3,4,5,6,7',\
		'07:30:00',\
		'23:00:00',\
		'1',\
		'07:30:00',\
		'10:00:00',\
		'1',\
		'admin',\
		'2017-11-09 11:39:05',\
		'admin',\
		'2017-11-09 12:23:57',\
		'N');".format(mastercode[0])
        sqldelivery="INSERT INTO `ddos_misc`.`master_store_t_delivery` (GUID ,\
        `STORE_CODE`,\
        `IS_DELIVERY_SPREAD`,\
        `DELIVERY_TIME_START`,\
        `DELIVERY_TIME_END`,\
        `DELIVERY_PRICE`,\
        `DELIVERY_AREA`,\
        `DELIVERY_EACH_TIME`,\
        `DELIVERY_STANDARD`,\
        `DELIVERY_FREE_PRICE`,\
        `DELIVERY_CHARGE_STANDARD`,\
        `DELIVERY_TYPE`,\
        `DELIVERY_BUSINESS`,\
        `IS_DELIVERY_DISTRICT`,\
        `BRAND_CODE`,\
        `CREATE_ID`,\
        `CREATE_TIME`,\
        `UPDATE_ID`,\
        `UPDATE_TIME`,\
        `DEL_FLAG`)\
    VALUES\
        (UUID(),\
            '{}',\
            '1',\
            '08:30:00',\
            '21:30:00',\
            '20',\
            '3',\
            '30',\
            '0',\
            '35',\
            '0',\
            '1',\
            NULL,\
            '1',\
            '1',\
            'admin',\
            '2017-11-09 11:39:27',\
            'admin',\
            '2017-11-09 12:24:07',\
            'N');".format(mastercode[0])
        sqlmarket="INSERT INTO `ddos_misc`.`master_store_t_org_tree` (GUID,\
        `STORE_CODE`,\
        `AREA_COMPANY`,\
        `SUB_COMPANY`,\
        `BUSINESS_CENTER`,\
        `AREA_SUPERVISOR_AREA`,\
        `AREA_SUPERVISOR`,\
        `AREA_SUPERVISOR_PHONE`,\
        `RGM`,\
        `RGM_PHONE`,\
        `STATE_FLAG`,\
        `BRAND_CODE`,\
        `CREATE_ID`,\
        `CREATE_TIME`,\
        `UPDATE_ID`,\
        `UPDATE_TIME`,\
        `DEL_FLAG`)\
    SELECT UUID(),{},a.AREACOMPANYID,\
            a.SUBCOMPANYID,\
        a.BUSINESSCENTERID,\
        'B001',\
             SUBSTRING(a.AREASUPERVISORAREAID,1,6),\
            '13540380437',\
            {},\
            NULL,\
            NULL,\
            '1',\
            'admin',\
            '2017-11-09 11:39:13',\
            'admin',\
            '2017-11-09 11:39:13',\
            'N'\
        FROM bpm_master as a  WHERE a.storecode={};".format(mastercode[0],mastercode[0],mastercode[1])
        sqluser="INSERT INTO `ddos_misc`.`master_fund_t_system_user` (\
        `guid`,\
        `user_id`,\
        `user_name`,\
        `password`,\
        `email`,\
        `mobile`,\
        `status`,\
        `create_time`,\
        `create_id`,\
        `update_time`,\
        `update_id`,\
        `del_flag`,\
        `area_code`,\
        `store_code`,\
        `itsm_create_user_flag`,\
        `user_type`,\
        `out_system_user_id`,\
        `uplevel_user_id`,\
        `brand_code`\
    )\
    VALUES\
        (\
            UUID(),\
            {},\
            {},\
            '8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918',\
            'fei_xiang@hoperun.com',\
            NULL,\
            '1',\
            '2017-11-09 12:49:24.000000',\
            'admin',\
            '2017-11-09 12:49:24.000000',\
            'admin',\
            'N',\
            NULL,\
            NULL,\
            '1',\
            '2',\
            NULL,\
            NULL,\
            '1'\
        );".format(mastercode[0],mastercode[0])
        print(sqlbaseinfo)
        print(sqlbussiness)
        print(sqldelivery)
        print(sqlmarket)
        print(sqluser)


        try:
                cursor.execute(sqlbaseinfo)
                cursor.execute(sqlbussiness)
                cursor.execute(sqldelivery)
                cursor.execute(sqlmarket)
                cursor.execute(sqluser)
        except Exception as e:
            conn.rollback()
            print('事务处理失败', e)
        else:
            conn.commit()  # 事务提交cursor.close()
            print("suess instert")
    cursor.close()
    conn.close
def instertchannel():
    conn = pymysql.connect(host="172.16.3.29"
                       , user="root",
                       password="dicos8888",
                       database="dicos_ios",
                       charset="utf8")
    cursor = conn.cursor()
    storecodes=[103190]
    list=[[1,1],
        [2,1],
        [3,2],
        [4,2],
        [5,2]]

    for storecode in storecodes:
        for channeltype in list:
            sql="INSERT INTO `dicos_ios`.`t_storechannel`\
        (`channel`, `storeCode`, `startTime`, `endTime`, `deliveryType`, `status`, `shutReason`, `shutMapTime`, `consignerId`, `consignerShowName`, `serviceFeeRate`)\
        VALUES ( '{}', '{}', '10:00', '22:00', '{}', '1', NULL, NULL, NULL, NULL,1);".format(channeltype[0],storecode,channeltype[1])
            print(sql)
            cursor.execute(sql)
    result = cursor.fetchall()
    if result != 0:
        print("插入channel成功")
    else:
        print("fail")

    conn.commit()
    cursor.close()
    conn.close()

# instertmaster()


def main():
    return_list=read_excel()
    print(return_list)
    instertmaster(return_list)



if __name__=="__main__":
    main()



