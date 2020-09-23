# Command line連Oralce資料庫產生Excel檔案

不需要安裝Office、Oracle Client, 只需要select、with開頭的SQL就可產生Excel, 可將SQL放至文字檔或TABLE內產生



## **環境設定**

1. 建立目錄 OracleToExcel
2. 將 OracleToExcel.exe、TNSNAMES.ORA、Oracle.ManagedDataAccess.dll、EPPlus.dll 放至此目錄
3. 設定 TNSNAMES.ORA



## **使用文字檔產生Excel**

1. 參數說明

   ![](images\1-1.jpg)

   

2. 建立一個文字檔內容為SELECT或是WITH的SQL

![](images\1-2.jpg)



3. 開啟COMMAND視窗執行

   ```shell
   D:\> OracleToExcel.exe hr/xxxxxx@orcl D:\emp D:\TEST.SQL
   ```

   

4. 產生結果

   ![](images\1-3.jpg)



## **使用TABLE產生Excel**

 1. 資料庫建立一個儲存SQL的TABLE

    ```sql
    CREATE TABLE SQL_TAB (
        SEQ            NUMBER(2) NOT NULL,
        SQL_SCRIPT     VARCHAR2(4000) NOT NULL,
        ASSIGN_VALUES  VARCHAR2(1000),
        LAST_DAT       DATE NOT NULL
    );
    
    ALTER TABLE SQL_TAB ADD (
        CONSTRAINT PK_SQL_TAB PRIMARY KEY (SEQ));
    
    CREATE OR REPLACE TRIGGER TRIG_SQL_TAB_BEF_CHG
    BEFORE UPDATE OR INSERT ON SQL_TAB
    REFERENCING OLD AS OLD NEW AS NEW
    FOR EACH ROW
    BEGIN
        :NEW.LAST_DAT := SYSDATE;
    END;
    ```

    

 2. 新增一筆測試資料

    | SEQ  | SQL_SCRIPT                                                   | ASSIGN_VALUES                    |
    | ---- | ------------------------------------------------------------ | -------------------------------- |
    | 1    | SELECT A.EMPLOYEE_ID 員工編號, A.FIRST_NAME, A.LAST_NAME, A.EMAIL, A.PHONE_NUMBER,<br/>    A.HIRE_DATE, A.JOB_ID, A.SALARY, A.DEPARTMENT_ID, B.DEPARTMENT_NAME<br/>FROM EMPLOYEES A, DEPARTMENTS B<br/>WHERE A.DEPARTMENT_ID = B.DEPARTMENT_ID | 1,6,S,文字測試<br />2,8,N,123456 |

    **欄位說明**

    ​    **SQL_SCRIPT :** SELECT或WITH開頭的SQL

    ​    **ASSIGN_VALUES 格式 :** 

    ​        `Excel列位置`,`Excel欄位置`,`S:字串、N:數值`,`顯示的內容`

    ​        `Excel列位置`,`Excel欄位置`,`S:字串、N:數值`,`顯示的內容`

    ​        ...

    

    ```sql
    INSERT INTO SQL_TAB(SEQ, SQL_SCRIPT, ASSIGN_VALUES)
    VALUES(1,
        'SELECT A.EMPLOYEE_ID 員工編號, A.FIRST_NAME, A.LAST_NAME, A.EMAIL, A.PHONE_NUMBER,' || CHR(10) ||
        '    A.HIRE_DATE, A.JOB_ID, A.SALARY, A.DEPARTMENT_ID, B.DEPARTMENT_NAME' || CHR(10) ||
        'FROM EMPLOYEES A, DEPARTMENTS B' || CHR(10) ||
        'WHERE A.DEPARTMENT_ID = B.DEPARTMENT_ID',
        '1,6,S,文字測試' || CHR(10) ||
        '2,8,N,123456'
    );
    COMMIT;
    ```

    

 3. 開啟COMMAND視窗執行

    ```shell
    D:\> OracleToExcel.exe hr/xxxxxx@orcl "D:\emp" "SELECT 2, 1, 1, SQL_SCRIPT, ASSIGN_VALUES FROM SQL_TAB WHERE SEQ = 1"
    ```

    

4. 產生結果

   ![](images\2-1.jpg)

   

   詳細參數說明如下

   

5. 指定保留空白列、欄

![](images\2-2.jpg)



6. 設定是否顯示欄位抬頭

![](images\2-3.jpg)

7. 指定欄位值

   ![](images\2-4.jpg)

   

8. 指定欄位值斷行處理

   

9. COMMAND LINE訊息說明

![](images\2-5.jpg)





