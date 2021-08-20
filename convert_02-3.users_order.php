<!-- 大專導入資料測試，★★★為須調整/確認參數 -->
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
$inputFileName = 'convert.excel/02.users-users_coupon-users_order.xlsx';

//透過套件功能來讀取 excel 檔
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);

//★★★取得指定名稱工作表
$worksheet = $spreadsheet->getSheetByName('users_order'); //或 $spreadsheet->->getSheet(2);

//讀取當前工作表(sheet)的資料列數
$highestRow = $worksheet->getHighestRow();

//依序讀取每一列，若是第一列為標題，建議 $i 從 2 開始
for($i = 2; $i <= $highestRow; $i++) {
    //若是某欄位值為空，代表那一列可能都沒資料，便跳出迴圈
    if( $worksheet->getCell('A'.$i)->getValue() === '' || 
        $worksheet->getCell('A'.$i)->getValue() === null ) break;
    
    //★★★讀取 cell 值
    $user_id = $worksheet->getCell('A'.$i)->getValue();
    $order_id = $worksheet->getCell('B'.$i)->getValue();
    $email = $worksheet->getCell('C'.$i)->getValue();
    $transport_area = $worksheet->getCell('D'.$i)->getValue();
    $transport_type = $worksheet->getCell('E'.$i)->getValue();
    $transport_payment = $worksheet->getCell('F'.$i)->getValue();
    $transport_arrival_time = $worksheet->getCell('G'.$i)->getValue();
    $recipient_email = $worksheet->getCell('H'.$i)->getValue();
    $recipient_name = $worksheet->getCell('I'.$i)->getValue();
    $recipient_phone_number = $worksheet->getCell('J'.$i)->getValue();
    $recipient_address = $worksheet->getCell('K'.$i)->getValue();

    $recipient_comments = $worksheet->getCell('L'.$i)->getValue();
    $invoice_type = $worksheet->getCell('M'.$i)->getValue();
    $invoice_carrier = $worksheet->getCell('N'.$i)->getValue();
    $invoice_carrier_number = $worksheet->getCell('O'.$i)->getValue();
    $coupon_code = $worksheet->getCell('P'.$i)->getValue();
    $card_number = $worksheet->getCell('Q'.$i)->getValue();
    $card_valid_date = $worksheet->getCell('R'.$i)->getValue();
    $card_ccv = $worksheet->getCell('S'.$i)->getValue();
    $card_holder = $worksheet->getCell('T'.$i)->getValue();
    $total = $worksheet->getCell('U'.$i)->getValue();
    $total_m = $worksheet->getCell('V'.$i)->getValue();
    
    //★★★寫入資料
    $sql = "INSERT INTO `users_order`(
        `user_id`, `order_id`, `email`,
        `transport_area`, `transport_type`, `transport_payment`,
        `transport_arrival_time`, `recipient_email`, `recipient_name`,
        `recipient_phone_number`, `recipient_address`,

        `recipient_comments`, `invoice_type`, `invoice_carrier`,
        `invoice_carrier_number`, `coupon_code`, `card_number`,
        `card_valid_date`, `card_ccv`, `card_holder`,
        `total`, `total_m`
        ) VALUES (
            '{$user_id}', '{$order_id}', '{$email}',
            '{$transport_area}', '{$transport_type}', '{$transport_payment}',
            '{$transport_arrival_time}', '{$recipient_email}', '{$recipient_name}',
            '{$recipient_phone_number}', '{$recipient_address}',

            '{$recipient_comments}', '{$invoice_type}', '{$invoice_carrier}',
            '{$invoice_carrier_number}', '{$coupon_code}', '{$card_number}',
            '{$card_valid_date}', '{$card_ccv}', '{$card_holder}',
            '{$total}', '{$total_m}'
            )";
    $stmt = $pdo->query($sql);

    //若是成功寫入資料
    if( $stmt->rowCount() > 0 ){
        //印出 AutoIncrement 的流水號碼 (若沒設定，預設為 0)
        echo $pdo->lastInsertId() . "\n";
    }
}