#include <stdio.h>
#include <windows.h>
#include <time.h>
#include <sql.h>
#include <sqlext.h>
#include <string.h>
#include "xlsxwriter.h"

#define DB_CFG_PATH "cfg/db.cfg"
#define CFG_PATH "cfg/paths.cfg"
#define LOG_PATH "logs/autoreports"
#define ZABBIX_CFG_PATH "cfg/zabbix.cfg"

#define LIGHT_GRAY 0xDDDDDD
#define NAME_MAX_LENGTH 100

typedef struct user{
    int user_id;
    char name[NAME_MAX_LENGTH];
    int dept_id;
} user;
typedef struct dept{
    int dept_id;
    char dept_name[NAME_MAX_LENGTH];
} dept;
typedef struct check_rec{
    int user_id;
    TIMESTAMP_STRUCT check_time;
} check_rec;
typedef struct dept_file_path{
    char dept_name[NAME_MAX_LENGTH];
    char file_path[260];
    struct dept_file_path *next;
} dept_file_path;

static FILE *log_file;

static int zabbix_enabled;
static char zabbix_ip[30];
static char zabbix_host[150];
static char zabbix_item[150];

static const char error_log[] = "ERROR";
static const char info_log[] = "INFO";

static int year;
static int month;
static int day;
const char *wdays[] = {"воск", "пон", "втор", "сред",
                          "четв", "пятн", "субб"};
const char *months[] = {"январь", "февраль", "март", "апрель",
                        "май", "июнь", "июль", "август",
                        "сентябрь", "октябрь", "ноябрь", "декабрь"};

void print_log(const char *type, const char *fmt, ...) {
    va_list vl;
    va_start(vl, fmt);

    time_t now = time(NULL);

    struct tm tm_now;
    localtime_s(&tm_now, &now);

    char message[1024];
    va_list vl_copy;
    va_copy(vl_copy, vl);
    vsnprintf(message, sizeof(message), fmt, vl_copy);
    va_end(vl_copy);

    if (zabbix_enabled && type == error_log) {
        char send_to_zabbix[1200];
        int n = snprintf(send_to_zabbix, sizeof(send_to_zabbix),
                         "zabbix_sender -z %s -s \"%s\" -k %s -o \"ERROR: %s\"",
                         zabbix_ip, zabbix_host, zabbix_item, message);
        if (n > 0 && n < (int)sizeof(send_to_zabbix)) {
            system(send_to_zabbix);
        } else {
            fprintf(stderr, "zabbix command too long or formatting error\n");
        }
    }

    if (log_file) {
        fprintf(log_file, "[%02d:%02d:%02d] %s: %s",
                tm_now.tm_hour, tm_now.tm_min, tm_now.tm_sec, type, message);
        fflush(log_file);
    }

    va_end(vl);
}
LONG WINAPI CrashHandler(EXCEPTION_POINTERS* ExceptionInfo)
{
    print_log(error_log, "Crash detected! Code: 0x%08X\nFault address: %p\n",
           ExceptionInfo->ExceptionRecord->ExceptionCode, ExceptionInfo->ExceptionRecord->ExceptionAddress);
    fclose(log_file);
    return EXCEPTION_EXECUTE_HANDLER;
}
void read_db_config(SQLCHAR *conn_str){
    FILE *file = fopen(DB_CFG_PATH, "r");
    if(!file){
        print_log(error_log, "failed to open database config %s\n\n", DB_CFG_PATH);
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
void read_zabbix_config(){
    FILE *file = fopen(ZABBIX_CFG_PATH, "r");
    if(!file){
        print_log(error_log, "failed to open zabbix config %s\n\n", ZABBIX_CFG_PATH);
        return;
    }
    int ch, i = 0;
    char tmp_str[50];
    const char *keys[] = {"enabled", "server", "host", "download.item"};
    char *values[] = {zabbix_ip, zabbix_host, zabbix_item};
    char *value = NULL;

    while((ch = fgetc(file)) != EOF){
        switch(ch){
            case '=':
                tmp_str[i] = '\0';
                i = 0;
                int j = 0;
                if(strcmp(tmp_str, keys[0]) == 0){
                    if((ch = fgetc(file)) != EOF && ch >= 48 && ch <= 57){
                        zabbix_enabled = ch-48;
                    }
                    if(!zabbix_enabled){
                        return;
                    }
                    break;
                }
                for(int k = 1; k < 4; k++){
                    if(strcmp(tmp_str, keys[k]) == 0){
                        value = values[k-1];
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
                i = 0;
                break;

            default:
                tmp_str[i] = ch;
                i++;
                break;
        }
    }
}
int free_dept_file_path(dept_file_path *file_path){
    if (!file_path) return 1;
    dept_file_path *cur = file_path;
    while (cur) {
        dept_file_path *next = cur->next;
        free(cur);
        cur = next;
    }
    return 0;
}

void set_columns_width(lxw_worksheet *worksheet){
    double columns_width[11] = {13.29, 8.29, 4.71, 20.71, 21.29, 0.86,
                                8.29, 2.71, 6.29, 8.29, 9.14};
    for(int i = 0; i<11; i++){
        worksheet_set_column(worksheet, i, i, columns_width[i], NULL);
    }
}

void sort_by_dept_id(user *users, int users_count){
    user tmp;
    for(int i = 0; i<users_count; i++){
        for(int j = 1; j<users_count-i; j++){
            if(users[j].dept_id < users[j-1].dept_id){
                tmp = users[j];
                users[j] = users[j-1];
                users[j-1] = tmp;
            }
        }
    }
}

void cp1251_to_utf8(const char *str, char *utf8){
    wchar_t wide[256];

    MultiByteToWideChar(1251, 0, str, -1, wide, 256);
    WideCharToMultiByte(CP_UTF8, 0, wide, -1, utf8, 512, NULL, NULL);
}
int utf8_to_cp1251(const char *utf8, char *cp1251, size_t cp1251_size) {
    wchar_t wide[512];

    int len = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wide, sizeof(wide)/sizeof(wide[0]));
    if (len == 0) {
        return 1;
    }

    len = WideCharToMultiByte(1251, 0, wide, -1, cp1251, (int)cp1251_size, NULL, NULL);
    if (len == 0) {
        return 2;
    }

    return 0;
}

int get_db_length(SQLINTEGER *db_length, const char *db_name, SQLHSTMT *hStmt, const char *condition){
    const char *prefix = "SELECT COUNT(*) FROM ";
    size_t len = strlen(prefix) + strlen(db_name) + 2;
    if(condition != NULL){
        len += strlen(condition);
    }

    char *sql_request = malloc(len);
    if(!sql_request){
        print_log(error_log, "Failed to allocate memory for SQL query\n");
        return 1;
    }

    strcpy(sql_request, prefix);
    strcat(sql_request, db_name);
    if(condition != NULL){
        strcat(sql_request, " ");
        strcat(sql_request, condition);
    }

    SQLRETURN ret = SQLExecDirect(hStmt, (SQLCHAR*)sql_request, SQL_NTS);

    if(!SQL_SUCCEEDED(ret)){
        print_log(error_log, "Error executing SQL query, request: %s\n", sql_request);
        free(sql_request);
        return 1;
    }

    SQLFetch(hStmt);
    SQLGetData(hStmt, 1, SQL_C_SLONG, db_length, 0, NULL);
    SQLCloseCursor(hStmt);

    free(sql_request);
    if(!*db_length){
        print_log(error_log, "Failed to get count of elements from db: %s\n", db_name);
        return 1;
    }
    return 0;
}
void get_users(SQLHSTMT *hStmt, user *users, int users_count){
    SQLExecDirect(hStmt, (SQLCHAR*)"SELECT Userid, Name, Deptid FROM dbo.Userinfo;", SQL_NTS);

    char tmp_user_id[20];
    char tmp_name[NAME_MAX_LENGTH];
    int tmp_dept_id;

    int current_user = 0;
    while(current_user < users_count && SQLFetch(hStmt) == SQL_SUCCESS){
        SQLGetData(hStmt, 1, SQL_C_CHAR, tmp_user_id, sizeof(tmp_user_id), NULL);
        SQLGetData(hStmt, 2, SQL_C_CHAR, tmp_name, sizeof(tmp_name), NULL);
        SQLGetData(hStmt, 3, SQL_C_LONG, &tmp_dept_id, 0, NULL);

        users[current_user].user_id = atoi(tmp_user_id);
        strcpy(users[current_user].name, tmp_name);
        users[current_user].dept_id = tmp_dept_id;

        current_user++;
    }
    sort_by_dept_id(users, users_count);
    SQLCloseCursor(hStmt);
}
void get_depts(SQLHSTMT *hStmt, dept *depts, int depts_count){
    SQLExecDirect(hStmt, (SQLCHAR*)"SELECT Deptid, DeptName FROM Dept;", SQL_NTS);

    int tmp_dept_id;
    char tmp_dept_name[NAME_MAX_LENGTH];

    int current_dept = 0;
    while(current_dept < depts_count && SQLFetch(hStmt) == SQL_SUCCESS){
        SQLGetData(hStmt, 1, SQL_C_LONG, &tmp_dept_id, 0, NULL);
        SQLGetData(hStmt, 2, SQL_C_CHAR, tmp_dept_name, sizeof(tmp_dept_name), NULL);

        depts[current_dept].dept_id = tmp_dept_id;
        strcpy(depts[current_dept].dept_name, tmp_dept_name);
        current_dept++;
    }
    SQLCloseCursor(hStmt);
}
void sort_rec(check_rec *rec, int size){
    check_rec tmp;
    for(int i = 0; i<size; i++){
        for(int j = 1; j < size-i; j++){
            if(rec[j].user_id < rec[j-1].user_id){
                tmp = rec[j];
                rec[j] = rec[j-1];
                rec[j-1] = tmp;
            }
        }
    }
}
void get_check_rec(SQLHSTMT *hStmt, check_rec *rec, int size){
    char sql_query[150];
    sprintf(sql_query,
     "SELECT Userid, CheckTime FROM dbo.Checkinout WHERE CheckTime BETWEEN '%04d-%02d-01T00:00:00.000' AND '%04d-%02d-%02dT23:59:59.999';",
    year, month, year, month, day);
    SQLExecDirect(hStmt, (SQLCHAR*)sql_query, SQL_NTS);
    
    char tmp_user_id[20];
    TIMESTAMP_STRUCT tmp_time;

    int current_rec = 0;
    while(current_rec < size && SQLFetch(hStmt) == SQL_SUCCESS){
        SQLGetData(hStmt, 1, SQL_C_CHAR, tmp_user_id, sizeof(tmp_user_id), NULL);
        SQLGetData(hStmt, 2, SQL_C_TYPE_TIMESTAMP, &tmp_time, 0, NULL);

        rec[current_rec].user_id = atoi(tmp_user_id);
        rec[current_rec].check_time = tmp_time;
        current_rec++;
    }
    sort_rec(rec, size);
    SQLCloseCursor(hStmt);
}
int find_check_rec(check_rec *rec, int size, int user_id){
    int lo = 0, hi = size-1;
    int mid, index = -1;
    while(lo <= hi){
        mid = (hi-lo)/2+lo;
        if(rec[mid].user_id > user_id){
            hi = mid-1;
        }else if(rec[mid].user_id < user_id){
            lo = mid+1;
        }else{
            index = mid;
            break;
        }
    }
    if(index == -1){
        return index;
    }
    while(rec[index].user_id == user_id){
        if(index == 0){
            return index;
        }
        index--;
    }
    return index+1;
}
int find_dept(dept *depts, int depts_count, int dept_id){
    int lo = 0, hi = depts_count-1;
    int mid;
    while(lo <= hi){
        mid = (hi-lo)/2+lo;
        if(depts[mid].dept_id > dept_id){
            hi = mid-1;
        }else if(depts[mid].dept_id < dept_id){
            lo = mid+1;
        }else{
            return mid;
        }
    }
    return -1;
}
void add_format(lxw_workbook *workbook, lxw_format **employe, lxw_format **check, lxw_format **time){
    *employe = workbook_add_format(workbook);

    format_set_font_name(*employe, "Segoe UI");
    format_set_font_size(*employe, 8);
    format_set_align(*employe, LXW_ALIGN_VERTICAL_CENTER);
    format_set_border(*employe, LXW_BORDER_THIN);
    format_set_border_color(*employe, LXW_COLOR_BLACK);
    format_set_bold(*employe);
    format_set_bg_color(*employe, LIGHT_GRAY);
    format_set_pattern(*employe, LXW_PATTERN_SOLID);
    format_set_text_wrap(*employe);

    *check = workbook_add_format(workbook);
    format_set_font_name(*check, "Segoe UI");
    format_set_font_size(*check, 8);
    format_set_align(*check, LXW_ALIGN_VERTICAL_CENTER);
    format_set_align(*check, LXW_ALIGN_CENTER);
    format_set_border(*check, LXW_BORDER_THIN);
    format_set_border_color(*check, LXW_COLOR_BLACK);
    format_set_text_wrap(*check);

    *time = workbook_add_format(workbook);
    format_set_font_name(*time, "Segoe UI");
    format_set_font_size(*time, 8);
    format_set_align(*time, LXW_ALIGN_VERTICAL_CENTER);
    format_set_border(*time, LXW_BORDER_THIN);
    format_set_border_color(*time, LXW_COLOR_BLACK);
    format_set_text_wrap(*time);
}
int create_dir_recursive(const char *path) {
    char tmp[MAX_PATH];
    char *p = NULL;
    size_t len;

    snprintf(tmp, sizeof(tmp), "%s", path);
    len = strlen(tmp);

    if (tmp[len - 1] == '\\')
        tmp[len - 1] = 0;

    for (p = tmp + 3; *p; p++) {
        if (*p == '\\') {
            *p = 0;
            CreateDirectoryA(tmp, NULL);
            *p = '\\';
        }
    }

    if (!CreateDirectoryA(tmp, NULL)) {
        DWORD err = GetLastError();
        if (err != ERROR_ALREADY_EXISTS)
            return 1;
    }

    return 0;
}
int read_files_paths(dept_file_path *first){
    FILE *file = fopen(CFG_PATH, "r");
    if(!file){
        print_log(error_log, "failed to open file %s\n", CFG_PATH);
        return 1;
    }
    dept_file_path *tmp = first;

    int slash_count = 0;
    int is_first = 1;
    int ch;
    int i, j;
    int is_key = 0;
    char tmp_str[MAX_PATH];
    char param[10] = {0};
    while((ch = fgetc(file)) != EOF){
        if(ch == ':'){
            is_key = 1;
        }else if(ch == '\"'){
            i = 0, j = 0; 
            while((ch = fgetc(file)) != EOF && ch != '\"'){
                tmp_str[i] = ch;
                if(ch == '$'){
                    param[j] = ch;
                    j++;
                }else if(strlen(param) > 0){
                    param[j] = ch;
                    j++;
                    if(strcmp(param, "${year}") == 0){
                        sprintf(param, "%d", year);
                        for(int k = i-j+1; k <= i; k++){
                            tmp_str[k] = 0;
                        }
                        strcat(tmp_str, param);
                        strcpy(param, "");
                        i = i-j+4;
                        for(int k = 0; k<10; k++){
                            param[k] = 0;
                        }
                    }
                }
                i++;
            }
            tmp_str[i] = 0;
            if(is_key){
                strcpy(tmp->file_path, tmp_str);
                is_key = 0;
            }else if(strlen(tmp_str) <= NAME_MAX_LENGTH){
                if(!is_first){
                    tmp->next = malloc(sizeof(dept_file_path));
                    tmp = tmp->next;
                }else{
                    is_first = 0;
                }
                strcpy(tmp->dept_name, tmp_str);
            }else{
                free_dept_file_path(first);
                print_log(error_log, "invalid %s\n", CFG_PATH);
                return 1;
            }
        }else if(ch == '/'){
            slash_count++;
            if(slash_count == 2){
                slash_count = 0;
                while((ch = fgetc(file)) != EOF && ch != '\n');
            }
        }
    }
    tmp->next = NULL;
    fclose(file);
    return 0;
}
int find_path_by_dept(const char *dept, char *path, dept_file_path *files_paths){
    dept_file_path *cur = files_paths;
    while(cur != NULL){
        if(strcmp(dept, cur->dept_name) == 0){
            strcpy(path, cur->file_path);
            return 0;
        }
        cur = cur->next;
    }
    return 1;
}
void write_user_to_xlsx(lxw_worksheet *worksheet, lxw_format *employe, lxw_format *check, lxw_format *time, user *user, check_rec *rec, int rec_count, const char *dept_name, int *current_row){
    char utf8[512];
    char tmp_str[256];
    int work_time_h;
    int work_time_m;
    int work_time_s;

    int day_check_count;
    TIMESTAMP_STRUCT first_check;
    int rec_index; 
    struct tm tmp_t = {0};
    tmp_t.tm_year = year-1900;
    tmp_t.tm_mon = month-1;
    
    worksheet_set_row(worksheet, *current_row, 15.4, NULL);
    worksheet_write_number(worksheet, *current_row, 0, user->user_id, employe);
    
    cp1251_to_utf8(user->name, utf8);
    worksheet_merge_range(worksheet, *current_row, 1, *current_row, 4, utf8, employe);

    worksheet_merge_range(worksheet, *current_row, 5, *current_row, 9, dept_name, employe);
        
    rec_index = find_check_rec(rec, rec_count, user->user_id);

    (*current_row)++;
    for(int j = 1; j <= day; j++){
        worksheet_set_row(worksheet, *current_row, 15.4, NULL);

        sprintf(tmp_str, "%02d.%02d.%04d", j, month, year);
        cp1251_to_utf8(tmp_str, utf8);
        worksheet_write_string(worksheet, *current_row, 0, utf8, check);

        tmp_t.tm_mday = j;
        mktime(&tmp_t);
        worksheet_write_string(worksheet, *current_row, 1, wdays[tmp_t.tm_wday], check);

        strcpy(utf8, "");
        day_check_count = 0;
        if(rec_index != -1){
            first_check = rec[rec_index].check_time;
            while(rec[rec_index].check_time.day == j && rec[rec_index].user_id == user->user_id){
                strcpy(tmp_str, "");
                sprintf(tmp_str, "%02d:%02d:%02d",
                        rec[rec_index].check_time.hour,
                        rec[rec_index].check_time.minute,
                        rec[rec_index].check_time.second);
                if(strlen(utf8) != 0){
                    strcat(utf8, " - ");
                }
                strcat(utf8, tmp_str);
                rec_index++;
                day_check_count++;
            }
        }
        worksheet_merge_range(worksheet, *current_row, 2, *current_row, 6, utf8, time);
        strcpy(utf8, "");
        if(day_check_count > 1){
            work_time_m = 0;
            work_time_h = 0;
            work_time_s = rec[rec_index-1].check_time.second - first_check.second;
            if(work_time_s < 0){
                work_time_s += 60;
                work_time_m--;
            }
            work_time_m += rec[rec_index-1].check_time.minute - first_check.minute;
            if(work_time_m < 0){
                work_time_m += 60;
                work_time_h--;
            }
            work_time_h += rec[rec_index-1].check_time.hour - first_check.hour;
            sprintf(utf8, "%02d:%02d:%02d", work_time_h, work_time_m, work_time_s);
        }
        worksheet_merge_range(worksheet, *current_row, 7, *current_row, 8, "", check);
        if(day_check_count > 0){
            worksheet_write_number(worksheet, *current_row, 7, day_check_count, check);
        }

        worksheet_write_string(worksheet, *current_row, 9, utf8, check);
        (*current_row)++;
    }
}

int main(int argc, char * argv[]){
    time_t now = time(NULL);
    struct tm *tm_now = localtime(&now);
    year = tm_now->tm_year+1900;
    month = tm_now->tm_mon+1;
    day = tm_now->tm_mday;
    
    zabbix_enabled = 0;

    char *log_path = malloc(sizeof(char)*100);
    sprintf(log_path, "%s/%02d-%02d-%04d.log", LOG_PATH, day, month, year);
    log_file = fopen(log_path, "a");
    free(log_path);
    print_log(info_log, "The program has started\n");
    SetUnhandledExceptionFilter(CrashHandler);

    read_zabbix_config();

    if(argc > 1){
        if(sscanf(argv[1], "%02d:%02d:%04d", &day, &month, &year) != 3){
            print_log(error_log, "The data from the arguments was read incorrectly. The date format must be DD:MM:YYYY.\n");
            print_log(info_log, "The program ended\n\n");
            fclose(log_file);
            return 1;
        }
    }

    dept_file_path *files_paths = malloc(sizeof(dept_file_path));
    files_paths->next = NULL;
    if(read_files_paths(files_paths)){
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
        print_log(error_log, "Database connection error\n");
        print_log(info_log, "The program ended\n\n");
        fclose(log_file);
        return 1;
    }
    SQLAllocHandle(SQL_HANDLE_STMT, hDbc, &hStmt);

    SQLINTEGER users_count = 0;
    if(get_db_length(&users_count, "dbo.Userinfo", hStmt, NULL)){
        print_log(error_log, "Failed to get the number of users\n");
        print_log(info_log, "The program ended\n\n");
        fclose(log_file);
        return 1;
    }
    user *users = malloc(users_count*sizeof(user));
    get_users(hStmt, users, users_count);

    SQLINTEGER depts_count = 0;
    if(get_db_length(&depts_count, "dbo.Dept", hStmt, NULL)){
        print_log(error_log, "Failed to get the number of depts\n");
        print_log(info_log, "The program ended\n\n");
        fclose(log_file);
        return 1;
    }
    dept *depts = malloc(depts_count*sizeof(dept));
    get_depts(hStmt, depts, depts_count);

    SQLINTEGER rec_count = 0;
    char condition[100];
    sprintf(condition, "WHERE CheckTime BETWEEN '%04d-%02d-01T00:00:00.000' AND '%04d-%02d-%02dT23:59:59.999'",
            year, month, year, month, day);
    if(get_db_length(&rec_count, "dbo.Checkinout", hStmt, condition)){
        print_log(error_log, "Failed to get the number of records\n");
        print_log(info_log, "The program ended\n\n");
        fclose(log_file);
        return 1;
    }
    check_rec *rec = malloc(rec_count*sizeof(check_rec));
    get_check_rec(hStmt, rec, rec_count);

    SQLFreeHandle(SQL_HANDLE_DBC, hDbc);
    SQLFreeHandle(SQL_HANDLE_ENV, hEnv);

    char utf8[512];
    char tmp_str[256];
    char current_dept_name[NAME_MAX_LENGTH];
    int current_dept = -1;
    int dept_index;
    int current_row = 0;

    lxw_format *employe; 
    lxw_format *check;
    lxw_format *time;

    lxw_workbook  *workbook = 0;
    lxw_worksheet *worksheet = 0;

    for(int i = 0; i < users_count; i++){
        if(current_dept != users[i].dept_id){
            if(workbook){
                lxw_error error = workbook_close(workbook);
                if(error){
                    print_log(error_log, "An error occurred while closing the workbook %s: %d\n", current_dept_name, error);
                }else{
                    print_log(info_log, "A report on the department was created: %s\n", current_dept_name);
                }
                current_row = 0;
            }
            current_dept = users[i].dept_id;
            workbook = 0;
            dept_index = find_dept(depts, depts_count, current_dept);
            if(dept_index != -1){
                cp1251_to_utf8(depts[dept_index].dept_name, current_dept_name);
                if(find_path_by_dept(current_dept_name, tmp_str, files_paths)){
                    continue;
                }
                utf8_to_cp1251(tmp_str, tmp_str, sizeof(tmp_str));
                create_dir_recursive(tmp_str);
                sprintf(utf8, "\\%02d %s.xlsx", month, months[month-1]);
                utf8_to_cp1251(utf8, utf8, sizeof(utf8));
                strcat(tmp_str, utf8);
                workbook  = workbook_new(tmp_str);
                worksheet = workbook_add_worksheet(workbook, current_dept_name);
                set_columns_width(worksheet);

                add_format(workbook, &employe, &check, &time);

                worksheet_set_row(worksheet, current_row, 30.75, NULL);
                worksheet_write_string(worksheet, current_row, 0, "Дата", check);
                worksheet_write_string(worksheet, current_row, 1, "Неделя", check);
                worksheet_merge_range(worksheet, current_row, 2, 0, 6, "Время отметки", time);
                worksheet_merge_range(worksheet, current_row, 7, 0, 8, "Раз", check);
                worksheet_write_string(worksheet, current_row, 9, "Время работы", check);

                current_row++;
                write_user_to_xlsx(worksheet, employe, check, time, &users[i], rec, rec_count, current_dept_name, &current_row);
            }
        }else if(workbook){
            write_user_to_xlsx(worksheet, employe, check, time, &users[i], rec, rec_count, current_dept_name, &current_row);
        }
    }
    if(workbook){
        lxw_error error = workbook_close(workbook);
        if(error){
            print_log(error_log, "An error occurred while closing the workbook %s: %d\n", current_dept_name, error);
        }else{
            print_log(info_log, "A report on the department was created: %s\n", current_dept_name);
        }
    }
    print_log(info_log, "The program ended\n\n");
    fclose(log_file);

    return 0;
}
