<?php

//upload.php

ini_set("memory_limit", "-1");
ini_set("display_startup_errors", 1);
ini_set("display_errors", 1);

include 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


if($_FILES["select_excel"]["name"] != "") {
    if ($_POST["beschrijving"] != "") {

        [$artnummerCol, $artnummerRow] = explode("-", $_POST["artnummer"]);
        [$beschrijvingCol, $beschrijvingRow] = explode("-", $_POST["beschrijving"]);
        [$hoeveelheidCol, $hoeveelheidRow] = explode("-", $_POST["hoeveelheid"]);
        [$eenheidCol, $eenheidRow] = explode("-", $_POST["eenheid"]);
        [$eenheidsprijsCol, $eenheidsprijsRow] = explode("-", $_POST["eenheidsprijs"]);

        if (isset($_POST["totaal"])) {
            [$totaalCol, $totaalRow] = explode("-", $_POST["totaal"]);
        }

        $spred = IOFactory::load($_FILES["select_excel"]["tmp_name"]);
        $shet = $spred->getSheetByName($_POST["sheet"]);

        $arr = [];

        for($i=1;$i<5000;$i++) {
            $row = $hoeveelheidRow + $i;

            $hoeveelheidVal = $shet->getCellByColumnAndRow($hoeveelheidCol + 1, $row)->getCalculatedValue();
            $eenheidsprijs = $shet->getCellByColumnAndRow($eenheidsprijsCol + 1, $row)->getCalculatedValue();


            // EX; A1:A6
            $range = PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($artnummerCol + 1) . $row . ":" . PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($beschrijvingCol + 1) . $row;

            $artnbeschrange = $shet->rangeToArray(
                $range,     // The worksheet range that we want to retrieve
                NULL,        // Value that should be returned for empty cells
                TRUE,        // Should formulas be calculated (the equivalent of getCalculatedValue() for each cell)
                FALSE,        // Should values be formatted (the equivalent of getFormattedValue() for each cell)
                FALSE       // Should the array be indexed by cell row and cell column
            );

            // Remove NULL from array
            $filteredArray = array_filter($artnbeschrange[0]);
            
            $artnummerBeschrijving = implode(" ", $filteredArray);
            // Eventueel  "&& count($filteredArray) > 1" toevoegen in de if
            if ($artnummerBeschrijving == "Opmerkingen")
            {
                break;
            }
            if (!empty($filteredArray)) {
                error_log($range);
                error_log($artnummerBeschrijving);

                        // check if "hoeveelheid" has a value
                if ($hoeveelheidVal != "" && $eenheidsprijs != "") {

                    // Row v/d Vorderingstaat
                    $curRow = count($arr) + 5;

                    $artnummer = $shet->getCellByColumnAndRow($artnummerCol + 1, $row)->getCalculatedValue();

                    // $eenheidsprijs = $shet->getCellByColumnAndRow($eenheidsprijsCol + 1, $row)->getCalculatedValue();

                    $totaal = "=IFERROR(\$B${curRow}*\$D${curRow},\"\")";
                    if (isset($totaalCol) && !is_numeric($shet->getCellByColumnAndRow($totaalCol + 1, $row)->getCalculatedValue())) {
                        // laat cell leeg wanneer bv: Vervalt of Ten laste bouwheer
                        $totaal = NULL;

                        // Vult Totaal in wanneer bv: Vervalt of Ten laste bouwheer
                        // $totaal = $shet->getCellByColumnAndRow($totaalCol + 1, $row)->getCalculatedValue();
                    }
            //  add value to the array
                    array_push($arr, [$artnummerBeschrijving, $hoeveelheidVal, $shet->getCellByColumnAndRow($eenheidCol + 1, $row)->getCalculatedValue(), $eenheidsprijs, $totaal,
                     0, "=\$C" . $curRow, "=IF(ISBLANK(OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),0,-2)),\"\",ROUND(OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),0,-2)*INDIRECT(ADDRESS(ROW(),COLUMN(D9000))),2))", "=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),,-3)", "=\$C" . $curRow, "=IF(ISBLANK(OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),0,-2)),\"\",ROUND(OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),0,-2)*INDIRECT(ADDRESS(ROW(),COLUMN(D9000))),2))", NULL, "=ROUND(INDIRECT(ADDRESS(ROW(),COLUMN(E9000)))-OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),,-2),2)"]);
                } else {
                    // Comment out to get only the values that have a quantity and a unitprice
                    array_push($arr, [$artnummerBeschrijving]);
                    continue;
                }

            }
        }

//        load spreadsheet
        $reader = IOFactory::createReader('Xlsx');
        $spreadsheet = $reader->load("template.xlsx");
        $worksheet = $spreadsheet->getActiveSheet();


//        TODO insert length of count($arr) columns before A6
        $worksheet->insertNewRowBefore(7, count($arr));

////        TODO add range of cells from array $arr to A5 $newsheet->fromArray($arr, NULL, "A5")
        $worksheet->fromArray($arr, null, "A5", true);

        $date = new DateTime('now');
        $enddate = $date->modify('last day of this month');
        
        $worksheet->setCellValue('F3', $date->format('01/m/Y') . ' t/m ' . $enddate->format('d/m/Y'));
        $worksheet->setCellValue('I3', 't/m ' . $enddate->format('d/m/Y'));

////        TODO save new document
        $downwriter = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $downwriter->save("result.xlsx");

        $message = '<script>window.location.href = "/result.xlsx"</script>';
        

    } else {

        $allowed_extension = array('xls', 'xlsx');
        $file_array = explode(".", $_FILES['select_excel']['name']);
        $file_extension = end($file_array);
        if (in_array($file_extension, $allowed_extension)) {
            $reader = IOFactory::createReader('Xlsx');
            if ($_POST["sheet"] != '') {
                $reader->setLoadSheetsOnly([$_POST["sheet"]]);
            } else {
                $reader->setLoadSheetsOnly(["Offerte klant"]);
            }
            $spreadsheet = $reader->load($_FILES['select_excel']['tmp_name']);
            $writer = IOFactory::createWriter($spreadsheet, 'Html');
            $message = $writer->save('php://output');
        } else {
            $message = '<div class="alert alert-danger">Only .xls or .xlsx file allowed</div>';
        }
    }
}
else
{
 $message = '<div class="alert alert-danger">Please Select File</div>';
}

echo $message;


?>
