#include <stdio.h>
#include <stdlib.h>
#include <windows.h>
#include <time.h>
#include <sql.h>
#include <sqlext.h>
#include <signal.h>
#include <stdarg.h>
#include "crosschex.h"

#define ERROR_LOG "ERROR"
#define INFO_LOG "INFO"

#define AUTO_REPORT_PATH "autoreports.exe"
#define CFG_PATH "cfg/devices.cfg"
#define LOG_PATH "logs/autodownload"
#define DB_CFG_PATH "cfg/db.cfg"

#define ZABBIX_IP "192.168.1.75"
#define ZABBIX_HOST "auto_crosschex"
#define ZABBIX_DATA_NAME "log.download.error"

#define MAX_IP_LEN 16
#define UTC_OFFSET 4

static FILE *log_file;

typedef struct{
    int sensor_id;
    char ip[MAX_IP_LEN];
} device_ip;

void print_log(char *type, char *str, ...){
    va_list vl;
    va_start(vl, str);

    time_t now = time(NULL);
    struct tm *tm_now = localtime(&now);

    if(0 == strcmp(type, ERROR_LOG)){
        char tmp[150];
        char send_to_zabbix[200];
        sprintf(tmp, "zabbix_sender -z %s -s \"%s\" -k %s -o \"ERROR: %s\"", ZABBIX_IP, ZABBIX_HOST, ZABBIX_DATA_NAME, str);
        vsprintf(send_to_zabbix, tmp, vl);
        system(send_to_zabbix);
    }

    fprintf(log_file, "[%02d:%02d:%02d] %s: ", tm_now->tm_hour, tm_now->tm_min, tm_now->tm_sec, type);
    vfprintf(log_file, str, vl);

    va_end(vl);
}
void signal_handler(int sig){
    print_log(ERROR_LOG, "Program terminated with signal %d\n\n", sig);
    fclose(log_file);
    exit(1);
}
void read_db_config(SQLCHAR *conn_str){
    FILE *file = fopen(DB_CFG_PATH, "r");
    if(!file){
        print_log(ERROR_LOG, "failed to open database config %s\n\n", DB_CFG_PATH);
        fclose(log_file);
        exit(1);
    }
    int ch, i = 0;
    char tmp_str[50];
    char server[50], db_name[50], uid[50], pwd[50];
    const char *keys[] = {"server", "database", "uid", "password"};
    char *values[] = {server, db_name, uid, pwd};
    char *value = NULL;


    while((ch = fgetc(file)) != EOF){
        switch(ch){
            case '=':
                tmp_str[i] = '\0';
                i = 0;
                int j = 0;
                for(int k = 0; k < 4; k++){
                    if(strcmp(tmp_str, keys[k]) == 0){
                        value = values[k];
                        break;
                    }
                }
                if(value){
                    while((ch = fgetc(file)) != EOF && ch != '\n'){
                        value[j] = ch;
                        j++;
                    }
                    value[j] = 0;
                    value = NULL;
                }
                break;
            
            case ' ':
                break;

            case '\n':
                break;

            default:
                tmp_str[i] = ch;
                i++;
                break;
        }
    }
    sprintf((char*)conn_str, "Driver={SQL Server};Server=%s\\SQLEXPRESS;Database=%s;UID=%s;PWD=%s;", server, db_name, uid, pwd);
}
int read_device_info(device_ip **devices_ip){
    FILE *file = fopen(CFG_PATH, "r");
    if(!file){
        print_log(ERROR_LOG, "Failed to open config file %s\n", CFG_PATH);
        return -1;
    }

    int slash_count = 0;
    int ch, i;
    int is_key = 0;
    int j = 0;
    char tmp_str[50];
    device_ip devices[50];

    while((ch = fgetc(file)) != EOF){
        if(ch == ':'){
            is_key = 1;
        }else if(ch == '\"'){
            i = 0; 
            while((ch = fgetc(file)) != EOF && ch != '\"'){
                tmp_str[i] = ch;
                i++;
            }
            tmp_str[i] = 0;
            if(is_key && strlen(tmp_str) < MAX_IP_LEN){
                strcpy(devices[j].ip, tmp_str);
                is_key = 0;
                j++;
            }else{
                devices[j].sensor_id = atoi(tmp_str);
            }
        }else if(ch == '/'){
            slash_count++;
            if(slash_count == 2){
                slash_count = 0;
                while((ch = fgetc(file)) != EOF && ch != '\n');
            }
        }
    }
    fclose(file);
    *devices_ip = malloc(sizeof(device_ip)*j);
    if (!*devices_ip) {
        print_log(ERROR_LOG, "Failed to allocate memory for devices_ip\n");
        return -1;
    }
    memcpy(*devices_ip, devices, sizeof(device_ip)*j);
    return j;
}

struct tm *sec_to_date(int seconds){
    struct tm start = {0};
    start.tm_year = 2000 - 1900;
    start.tm_mon  = 0;
    start.tm_mday = 2;
    start.tm_hour = 0;
    start.tm_min  = 0;
    start.tm_sec  = 0;

    time_t t = mktime(&start);

    t += seconds;

    return gmtime(&t);
}
void rev_bytes(unsigned char *bytes, int size){
    char tmp;
    for(int i = 0; i<size/2; i++){
        tmp = bytes[i];
        bytes[i] = bytes[size-1-i];
        bytes[size-1-i] = tmp;
    }
}
void insert_rec(SQLHSTMT *hStmt, unsigned int user_id, const struct tm *time, int sensor_id){
    SQLRETURN ret;
    char sql_query[150];

    sprintf(sql_query, "INSERT INTO dbo.Checkinout (Userid, CheckTime, Sensorid) VALUES ('%u', convert(datetime, '%04d-%02d-%02d %02d:%02d:%02d', 20), %d)",
            user_id, time->tm_year+1900, time->tm_mon+1, time->tm_mday, time->tm_hour+UTC_OFFSET, time->tm_min, time->tm_sec, sensor_id);

    ret = SQLExecDirect(hStmt, (SQLCHAR*)sql_query, SQL_NTS);
    if (!(ret == SQL_SUCCESS || ret == SQL_SUCCESS_WITH_INFO)) {
        print_log(ERROR_LOG, "An error occurred while executing a database query. Query: %s\n", sql_query);
    }
    SQLCloseCursor(hStmt);
}
int cmp_records(CCHEX_RET_RECORD_INFO_STRU *rec1, CCHEX_RET_RECORD_INFO_STRU *rec2){
    for(int i = 0; i < MAX_EMPLOYEE_ID; i++){
        if(rec1->EmployeeId[i] != rec2->EmployeeId[i]){
            return 0;
        }
    }
    for(int i = 0; i < 4; i++){
        if(rec1->Date[i] != rec2->Date[i]){
            return 0;
        }
    }
    return 1;
}

int main(){
    time_t now = time(NULL);
    struct tm *tm_now = localtime(&now);
    char log_path[64];
    
    snprintf(log_path, sizeof(log_path), "%s/%02d-%02d-%04d.log", LOG_PATH, tm_now->tm_mday, tm_now->tm_mon+1, tm_now->tm_year+1900);
    log_file = fopen(log_path, "a");
    print_log(INFO_LOG, "The program has started\n");

    signal(SIGSEGV, signal_handler);
    signal(SIGABRT, signal_handler);
    signal(SIGFPE,  signal_handler);

    device_ip *devices = NULL;
    int devices_len = read_device_info(&devices);
    if(!devices){
        print_log(ERROR_LOG, "Failed to get devices from cfg file\n");
        print_log(INFO_LOG, "The program ended\n\n");

        fclose(log_file);
        return 1;
    }

    SQLHENV hEnv;
    SQLHDBC hDbc;
    SQLHSTMT hStmt;
    SQLRETURN ret;

    SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &hEnv);
    SQLSetEnvAttr(hEnv, SQL_ATTR_ODBC_VERSION, (void *)SQL_OV_ODBC3, 0);
    SQLAllocHandle(SQL_HANDLE_DBC, hEnv, &hDbc);

    SQLCHAR *conn_str = malloc(sizeof(SQLCHAR)*150);
    read_db_config(conn_str);
    ret = SQLDriverConnect(hDbc, NULL, conn_str, SQL_NTS, NULL, 0, NULL, SQL_DRIVER_NOPROMPT);
    free(conn_str);
    
    if(!SQL_SUCCEEDED(ret)){
        print_log(ERROR_LOG, "Database connection error\n");
        print_log(INFO_LOG, "The program ended\n\n");

        fclose(log_file);
        return 1;
    }
    SQLAllocHandle(SQL_HANDLE_STMT, hDbc, &hStmt);

    void *handle = NULL;
    int devIdx;
    int type = 0;
    char buffer[1024];
    int len;
    CCHEX_RET_DEV_STATUS_STRU dev_info;
    CCHEX_RET_RECORD_INFO_STRU prev_ret = {0};

    CChex_Init();
    
    for(int i = 0; i<devices_len; i++){
        handle = CChex_Start();
        CChex_SetSdkConfig(handle, 0, 1, 0);

        dev_info.NewRecNum = -1;
        devIdx = CCHex_ClientConnect(handle, devices[i].ip, 5010);

        if(devIdx < 0){
            CChex_Stop(handle);
            continue;
        }

        Sleep(200);
        for(int j = 0; j<50; j++){
            len = CChex_Update(handle, &devIdx, &type, buffer, sizeof(buffer));
            if (len > 0) {
                if(type == 19){
                    memcpy(&dev_info, buffer, sizeof(dev_info));
                    if(dev_info.NewRecNum > 0){
                        CChex_DownloadAllNewRecords(handle, devIdx);
                    }
                    Sleep(200);
                    break;
                }
            }
            Sleep(200);
        }
        if(dev_info.NewRecNum == -1){
            print_log(ERROR_LOG, "Failed to get data from device number %d\n", devices[i].sensor_id);
            CChex_Stop(handle);
            Sleep(1000);
            continue;
        }
        print_log(INFO_LOG, "Number of new entries on device number %d: %u\n", devices[i].sensor_id, dev_info.NewRecNum);

        int j = 0;
        while(j < dev_info.NewRecNum){
            len = CChex_Update(handle, &devIdx, &type, buffer, sizeof(buffer));
            if(len > 0 && type == 71){
                CCHEX_RET_RECORD_INFO_STRU ret;
                memcpy(&ret, buffer, sizeof(ret));
                if(cmp_records(&ret, &prev_ret)){
                    Sleep(200);
                    continue;
                }
                memcpy(&prev_ret, &ret, sizeof(prev_ret));

                unsigned int seconds, employee_id;
                rev_bytes(ret.Date, 4);
                rev_bytes(ret.EmployeeId, MAX_EMPLOYEE_ID);

                memcpy(&seconds, ret.Date, sizeof(seconds));
                memcpy(&employee_id, ret.EmployeeId, sizeof(employee_id));
                struct tm *date = sec_to_date(seconds);
                
                insert_rec(hStmt, employee_id, date, devices[i].sensor_id);

                print_log(INFO_LOG, "A new recording has been downloaded, Employee ID: %u\n",employee_id);

                j++;
            }
            Sleep(200);
        }
        CCHex_ClientDisconnect(handle, devIdx);
        Sleep(2000);

        while (CChex_Update(handle, &devIdx, &type, buffer, sizeof(buffer)) > 0);
        CChex_Stop(handle);
    }

    SQLFreeHandle(SQL_HANDLE_DBC, hDbc);
    SQLFreeHandle(SQL_HANDLE_ENV, hEnv);

    system(AUTO_REPORT_PATH);

    print_log(INFO_LOG, "Auto report was launched\n");

    if(tm_now->tm_mday == 1){
        tm_now->tm_mday = 0;
        mktime(tm_now);
        char start_query[100];
        sprintf(start_query, "%s %02d:%02d:%04d", AUTO_REPORT_PATH, tm_now->tm_mday, tm_now->tm_mon+1, tm_now->tm_year+1900);
        system(start_query);

        print_log(INFO_LOG, "The auto report for the previous month was launched\n");
    }

    print_log(INFO_LOG, "The program ended\n\n");
    fclose(log_file);
    return 0;
}
