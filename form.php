<?php
header('Content-Type: text/html; charset=utf-8');

require_once 'conf/site.php';
require_once 'classes/PHPExcel.php';

$id = time();

$fileName = $id.'.xlsx';
$filePath = $siteConf['xlsxDir'].$fileName;


function cellColor($cells,$color){
    global $objPHPExcel;
    $objPHPExcel->getActiveSheet()->getStyle($cells)->getFill()
    ->applyFromArray(array('type' => PHPExcel_Style_Fill::FILL_SOLID,
    'startcolor' => array('rgb' => $color)
    ));
}

$objPHPExcel = new PHPExcel();
 
$objPHPExcel->getDefaultStyle()->getFont()
    ->setName('Arial')
    ->setSize(10);

$objPHPExcel->getDefaultStyle()->getAlignment()
        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP)
        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)
        ->setWrapText(true);


$objPHPExcel->setActiveSheetIndex(0);
$activeSheet = $objPHPExcel->getActiveSheet();


$headerStyleArray = array(
    'font'  => array(
        'size'  => 10,
        'bold'  => true
    )
);

$valueStyleArray = array(
    'font'  => array(
        'size'  => 9,
    )
);

$valueItalicStyleArray = array(
    'font'  => array(
        'size'  => 8,
        'italic' => true,
    )
);

$surname        = $_POST['surname'] ? $_POST['surname'] : '';
$name           = $_POST['name'] ? $_POST['name'] : '';
$middleName     = $_POST['middle-name'] ? $_POST['middle-name'] : '';
$nationality    = $_POST['nationality'] == 1 ? 'Гражданин РФ' : 'Не гражданин РФ';
$phoneFirst     = $_POST['phone-first'] ? $_POST['phone-first'] : '';
$phoneSecond    = $_POST['phone-second'] ? $_POST['phone-second'] : '';
$birthday       = $_POST['birthday-day'].'.'.$_POST['birthday-month'].'.'.$_POST['birthday-year'];
$birthday       = $birthday ? $birthday : '';
$region         = $_POST['region'] ? $_POST['region'] : '';
$service        = $_POST['service'] ? $_POST['service'] : '';
$purpose        = $_POST['purpose'] ? $_POST['purpose'] : '';
$amount         = $_POST['amount'] ? $_POST['amount'] : '';
$rent           = $_POST['rent'] ? $_POST['rent'] : '';
$downpayment    = $_POST['downpayment'] ? $_POST['downpayment'] : '';
$experienceNaim = $_POST['experience-naim'] ? $_POST['experience-naim'] : '';
$receipt        = $_POST['receipt'] ? $_POST['receipt'] : '';
$verificationNaim = $_POST['verification-naim'] ? $_POST['verification-naim'] : '';
$servises       = $_POST['servises'] ? $_POST['servises'] : '';
$marital        = $_POST['marital'] ? $_POST['marital'] : '';
$dop1           = $_POST['dop1'] ? 'есть' : 'нет';
$dop2           = $_POST['dop2'] ? 'не было(снята)' : 'есть';

$hearderCellArray = array(
    array('номер заявки', $id),
    array('Фамилия', $surname),
    array('Имя', $name),
    array('Отчество', $middleName),
    array('гражданство', $nationality),
    array('телефон 1', $phoneFirst),
    array('телефон 2', $phoneSecond),
    array('дата рождения', $birthday),
    array('Регион (область)', $region),
    array('интересна услуга', $service),
    array('цель кредита', $purpose),
    array('сумма кредита (руб)', $amount),
    array('сумма подтверждённого дохода', $rent),
    array('размер первоначального взноса (руб)', $downpayment),
    array('стаж на последнем месте работы', $experienceNaim),
    array('когда нужны деньги', $receipt),
    array('как будет подтверждать доход', $verificationNaim),
    array('дополнительные запросы', $servises),
    array('семейное положение', $marital),
    array('созаемщики', $dop1),
    array('судимость', $dop2),
    array('Комментарий оператора', ''),
);

$alphas = range('A', 'V');

$backgroundColor = 'FFD1D9';

foreach($alphas as $k => $alpha){
        
    if(isset($hearderCellArray[$k])){
        $activeSheet->setCellValue($alpha.'1', $hearderCellArray[$k][0]);
        $activeSheet->setCellValue($alpha.'2', $hearderCellArray[$k][1]);
    }

    if($alpha == 'O'){
        $backgroundColor = 'F5B916';
    }
    cellColor($alpha.'1', $backgroundColor);
    
    $activeSheet->getStyle($alpha.'1')->applyFromArray($headerStyleArray)->getAlignment()->setIndent(1);
    $activeSheet->getStyle($alpha.'2')->applyFromArray($valueStyleArray)->getAlignment()->setIndent(1); 
    $activeSheet->getColumnDimension($alpha)->setWidth(15);
}

$activeSheet->getColumnDimension('A')->setWidth(18);

$activeSheet->getStyle('L2')->applyFromArray(
    array(
        'font'  => array(
            'size'  => 9,
            'bold'  => true
        )
    )
);

$activeSheet->getStyle('J2:K2')->applyFromArray($valueItalicStyleArray);
$activeSheet->getStyle('O2:V2')->applyFromArray($valueItalicStyleArray);

$objPHPExcel->getActiveSheet()->setTitle('Заявка');

$objPHPExcel->setActiveSheetIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
   
$objWriter->save($filePath);


//* Отправляем на FTP

$ftpServer = $siteConf['remoteHost'];
$ftpUserName = $siteConf['remoteLogin'];
$ftpUserPass = $siteConf['remotePass'];


$serverFile = $siteConf['remoteDir'].$fileName;

if(file_exists($filePath)){

    $connId = ftp_connect($ftpServer);

    ftp_login($connId, $ftpUserName, $ftpUserPass);

    ftp_pasv($connId, true);
    
    ftp_put($connId, $serverFile, $filePath, FTP_BINARY);

    ftp_close($connId);
}