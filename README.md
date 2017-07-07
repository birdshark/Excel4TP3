###Excel4TP3
---
Excel4TP3是改造自[PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet "PhpSpreadsheet")的一个库,源库只支持php5.6及以上版本。
这里做了下改造,可以支持php5.4版本的,使用方法还是跟例子里面介绍的一样

这个是为适应ThinkPHP3框架改造的一个库

###使用方法
----
将Common/Config/config.php里的代码拷贝添加到自己项目的Common/Config/config.php里面

然后你可以在你的控制器里这样使用

		$spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);

        $styleThinBlackBorderOutline = [
            'borders' => [
                'outline' => [
                    'style' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];
        $spreadsheet->getActiveSheet()->getStyle('A1:E3')->applyFromArray($styleThinBlackBorderOutline);

        $sharedStyle2 = new \PhpOffice\PhpSpreadsheet\Style();
        $sharedStyle2->applyFromArray(
            [
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                ],
                'fill' => [
                    'type' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'color' => ['argb' => 'FFFFFF00'],
                ],
                'borders' => [
                    'bottom' => ['style' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN],
                    'right' => ['style' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM],
                ],
            ]
        );


        $spreadsheet->getActiveSheet()->mergeCells('A1:E1');
        $spreadsheet->getActiveSheet()->duplicateStyle($sharedStyle2, 'A1:E1');
        $spreadsheet->getActiveSheet()->getCell('A1')->setValue('数据汇总');

        $spreadsheet->getActiveSheet()
            ->setCellValue('A2', '非中文单词数')
            ->setCellValue('B2', '中朝字数')
            ->setCellValue('C2', '字数')
            ->setCellValue('D2', '字符数(不计空格)')
            ->setCellValue('E2', '稿件重复率');
        $spreadsheet->getActiveSheet()
            ->setCellValue('A3', $task['at_ww_count'])
            ->setCellValue('B3', $task['at_ew_count'])
            ->setCellValue('C3', $task['at_ew_count'] + $task['at_ww_count'])
            ->setCellValue('D3', $task['at_wcws_count'])
            ->setCellValue('E3', $task['at_repeat_rate']+"%");



        $spreadsheet->getActiveSheet()->mergeCells('A5:C5');
        $spreadsheet->getActiveSheet()->duplicateStyle($sharedStyle2, 'A5:C5');
        $spreadsheet->getActiveSheet()->getCell('A5')->setValue('数据详情');
        $spreadsheet->getActiveSheet()
            ->setCellValue('A6', '文档名称')
            ->setCellValue('B6', '非中文单词数')
            ->setCellValue('C6', '中朝字数');

        foreach ($details as $lk => $li) {
            $index = $lk + 7;
            $spreadsheet->getActiveSheet()
                ->setCellValue('A'.$index, $li['af_name'])
                ->setCellValue('B'.$index, $li['af_ww_count'])
                ->setCellValue('C'.$index, $li['af_ew_count']);
        }


        $title = $task['at_name'].'-数据表'.date('Y年m月d日H时i分');
        // Rename worksheet
        $spreadsheet->getActiveSheet()->setTitle('Operation table');

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);



        // Redirect output to a client’s web browser (Xlsx)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$title.'.xlsx"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
		// header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');

附上[Guide](http://phpspreadsheet.readthedocs.io/en/develop/ "Guide")

