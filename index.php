<?php
/* Incluir la libreria PHPExcel */
require_once './Classes/PHPExcel.php';

// Crea un nuevo objeto PHPExcel
$objPHPExcel = new PHPExcel();

// Establecer propiedades. No son totes precis
$objPHPExcel->getProperties()
->setCreator("VCatala")
->setLastModifiedBy("Vcatala")
->setTitle("Documento Excel de Prueba")
->setSubject("Documento Excel de Prueba")
->setDescription("Demostracion sobre como crear y abrir archivos de Excel desde PHP.")
->setKeywords("Excel Office 2007 openxml php")
->setCategory("Pruebas de Excel");

// Agregar Informacion. Tambe es pot clavar dins de un for, while, foreach...
$objPHPExcel->setActiveSheetIndex(0)
->setCellValue('A1', 'Ana aço es una prova del codi funcionant')
->setCellValue('A2', 'Valor 1')
->setCellValue('B2', 'Valor 2')
->setCellValue('C2', 'Total')
->setCellValue('A3', '10')
->setCellValue('B3', '7')
->setCellValue('C3', '=sum(A3:B3)');

// Renombrar Hoja. No es precis fer-ho
$objPHPExcel->getActiveSheet()->setTitle('Prueba TSI');

// Establecer la hoja activa, para que cuando se abra el documento se muestre primero. Si no heu poses no funciona.
$objPHPExcel->setActiveSheetIndex(0);

// Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel.
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="pruebaReal.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
?>
