<?
require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_before.php");
CModule::IncludeModule("iblock");
?>
<form name="form" action="" method="get">
От: <input type="text" value="" placeholder="2020-01-30" name="from" required />
До: <input type="text" value="" placeholder="2020-12-31" name="to" required />
<input type="hidden" name="do" value="send" />
<input type="submit" />
</form>
<?
if($_REQUEST['do']=='send'){
	Bitrix\Main\Loader::includeModule("sale");
	CModule::IncludeModule("iblock");
	CModule::IncludeModule("catalog");
	global $DB;
	$els = CIBlockElement::GetList(array(), array("IBLOCK_ID" => 3), false, false, array("NAME", "ID"));
	while($el = $els -> GetNext()){
		$_products[$el['NAME']]['COUNT'] = 0;
		$_products[$el['NAME']]['PRICE'] = 0;
	}
	$res = $DB->Query("SELECT * FROM b_sale_order_delivery WHERE DATE_DEDUCTED between '".$_REQUEST['from']." 00:00:00' AND '".$_REQUEST['to']." 23:59:59' AND DEDUCTED = 'Y'");
$price=0;
$pprice = 0;
	while($r = $res->GetNext()){
		$order = $DB->Query("SELECT b_sale_order_dlv_basket.BASKET_ID FROM b_sale_order_dlv_basket, b_sale_order_delivery WHERE b_sale_order_dlv_basket.ORDER_DELIVERY_ID = '".$r['ID']."' AND b_sale_order_delivery.ORDER_ID = '".$r['ORDER_ID']."' GROUP BY BASKET_ID");
		while($basket_item = $order -> GetNext()){
			$basket = $DB->Query("SELECT * FROM b_sale_basket WHERE ID = '".$basket_item['BASKET_ID']."'");
			$products[$r['ORDER_ID']] = array();
			while($ord = $basket->GetNext()){
				//$products[$r['ORDER_ID']][$ord['NAME']] = $ord['BASE_PRICE']*$ord['QUANTITY'];
				$price+=$ord['BASE_PRICE']*$ord['QUANTITY'];
				$ar_res = CCatalogProduct::GetByID($ord['PRODUCT_ID']);
				$pprice+=$ar_res['PURCHASING_PRICE']*$ord['QUANTITY'];
				$_products[$ord['NAME']]['PURCHASING_PRICE']+=$ar_res['PURCHASING_PRICE']*$ord['QUANTITY'];
				$_products[$ord['NAME']]['PRICE']+=$ord['PRICE']*$ord['QUANTITY'];
				$_products[$ord['NAME']]['COUNT']+=$ord['QUANTITY'];
			}
		}
	}
	

ksort($_products);

require_once __DIR__ . '/PHPExcel-1.8/Classes/PHPExcel.php';
require_once __DIR__ . '/PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';

	$xls  = new PHPExcel();
	
	$xls->setActiveSheetIndex(0);
// Получаем активный лист
	$sheet = $xls->getActiveSheet();
	$sheet->setTitle('Отчет по продажам');
	$sheet->setCellValue("A1", 'Товар');
	$sheet->setCellValue("B1", 'Количество');
	$sheet->setCellValue("C1", 'Цена');
	$sheet->setCellValue("D1", 'Себестоимость');
	$sheet->setCellValue("E1", 'Коэффициент');
	

	?>
	<table border="0">
	<tr><td width="50%">Товар</td><td>Количество, отгруженное за интервал</td><td>Общая сумма отгруженного товара</td><td>Закупочная цена</td><td>Коэффициент</td></tr>
	<? 
	$i=1;
	foreach($_products as $k => $v){
	if($v['COUNT']!=''){
		$i++;

			$sheet->setCellValue("A".$i, $k);
			$sheet->setCellValue("B".$i, $v['COUNT']);
			$sheet->setCellValue("C".$i, $v['PRICE']);
			$sheet->setCellValue("D".$i, $v['PURCHASING_PRICE']);
			if($v['PURCHASING_PRICE']>0){
			$sheet->setCellValue("E".$i, round($v['PRICE']/$v['PURCHASING_PRICE'], 2));
			}

	?>

	<tr><td width="50%"><?=$k;?></td><td><?=$v['COUNT'];?></td><td><?=$v['PRICE'];?></td><td><?=$v['PURCHASING_PRICE'];?></td><td><?if($v['PURCHASING_PRICE']>0){ echo round($v['PRICE']/$v['PURCHASING_PRICE'], 2);}?></td></tr>
	<?
	}
	}
		$i++;
	$data = "Итого отгружено заказов на сумму:".CurrencyFormat($price, "RUB").", общая сумма реализации: ".CurrencyFormat($pprice, "RUB");
	if($pprice>0){ $data.=" Коэффициент: ".round($price/$pprice, 2);}
		$sheet->setCellValue("A".$i, $data);

	
		$objWriter = new PHPExcel_Writer_Excel5($xls);
		$objWriter->save('report.xls');
	
	?>
	</table>
	<?
		echo $i;
		?>
	<hr />
	<b>Итого отгружено заказов на сумму: <?=CurrencyFormat($price, "RUB");?>, общая сумма реализации: <?=CurrencyFormat($pprice, "RUB");?>, коэффициент: <?if($pprice>0){ echo round($price/$pprice, 2);}?></b>
	<br /><a href="report.xls" target="_blank">Отчет в excell</a>
	<?

}  