<?php
//讀取 composer 所下載的套件
require_once('vendor/autoload.php');

//資料庫主機設定
const DRIVER_NAME = "mysql"; //使用哪一家資料庫
const DB_HOST = "localhost"; //資料庫網路位址 (127.0.0.1)
const DB_USERNAME = "root"; //資料庫連線帳號
const DB_PASSWORD = ""; //資料庫連線密碼
const DB_NAME = "techit"; //★★★指定資料庫
const DB_CHARSET = "utf8mb4"; //指定字元集，亦即是文字的編碼來源
const DB_COLLATE = "utf8mb4_unicode_ci"; //在字元集之下的排序依據

//資料庫連線變數
$pdo = null;

//錯誤處理
try {
    //設定 PDO 屬性 (Array 型態)
    $options = [
        PDO::ATTR_EMULATE_PREPARES => false,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION,
        PDO::MYSQL_ATTR_INIT_COMMAND => 'SET NAMES ' . DB_CHARSET . ' COLLATE ' . DB_COLLATE
    ];

    //PDO 連線
    $pdo = new PDO(
        DRIVER_NAME. ':host=' . DB_HOST . ';dbname=' . DB_NAME . ';charset=' .DB_CHARSET, 
        DB_USERNAME, 
        DB_PASSWORD, 
        $options
    );
} catch (PDOException $e) {
    echo "資料庫連結失敗，訊息: " . $e->getMessage();
    exit();
}


/**
 * 官方範例
 * URL: https://phpspreadsheet.readthedocs.io/en/latest/
 */


//★★★你的 excel 檔案路徑 (含檔名)
$inputFileName = 'convert.excel/02.users-users_coupon-users_follow.xlsx';

//透過套件功能來讀取 excel 檔
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);

//★★★取得指定名稱工作表
$worksheet = $spreadsheet->getSheetByName('users'); //或 $spreadsheet->->getSheet(2);

//讀取當前工作表(sheet)的資料列數
$highestRow = $worksheet->getHighestRow();

//依序讀取每一列，若是第一列為標題，建議 $i 從 2 開始
for($i = 2; $i <= $highestRow; $i++) {
    //若是某欄位值為空，代表那一列可能都沒資料，便跳出迴圈
    if( $worksheet->getCell('A'.$i)->getValue() === '' || 
        $worksheet->getCell('A'.$i)->getValue() === null ) break;
    
    //★★★讀取 cell 值
    $user_id = $worksheet->getCell('A'.$i)->getValue();
    $email = $worksheet->getCell('B'.$i)->getValue();
    $pwd = $worksheet->getCell('C'.$i)->getValue();
    $user_name = $worksheet->getCell('D'.$i)->getValue();
    $photo_sticker = $worksheet->getCell('E'.$i)->getValue();
    $phone_number = $worksheet->getCell('F'.$i)->getValue();
    $birthday = $worksheet->getCell('G'.$i)->getValue();
    $class = $worksheet->getCell('H'.$i)->getValue();
    $address = $worksheet->getCell('I'.$i)->getValue();
    $store_a = $worksheet->getCell('J'.$i)->getValue();
    $store_b = $worksheet->getCell('K'.$i)->getValue();
    
    //★★★寫入資料
    $sql = "INSERT INTO `users`(
        `user_id`,
        `email`, 
        `pwd`, 
        `user_name`,
        `photo_sticker`, 
        `phone_number`, 
        `birthday`,
        `class`, 
        `address`, 
        `store_a`,
        `store_b`
        ) VALUES (
            '{$user_id}', 
            '{$email}', 
            '{$pwd}', 
            '{$user_name}', 
            '{$photo_sticker}', 
            '{$phone_number}', 
            '{$birthday}', 
            '{$class}', 
            '{$address}', 
            '{$store_a}', 
            '{$store_b}'
            )";
            $stmt = $pdo->query($sql);

    //若是成功寫入資料
    if( $stmt->rowCount() > 0 ){
        //印出 AutoIncrement 的流水號碼 (若沒設定，預設為 0)
        echo $pdo->lastInsertId() . "\n";
    }
}