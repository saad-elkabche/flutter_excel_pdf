import 'dart:convert';

import 'dart:io';

import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:syncfusion_flutter_pdf/pdf.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as Excel;
import 'package:open_file/open_file.dart';
import 'package:universal_html/html.dart' show AnchorElement;







class Home extends StatefulWidget {
  const Home({Key? key}) : super(key: key);

  @override
  State<Home> createState() => _HomeState();
}

class _HomeState extends State<Home> {

  final Excel.Borders borders=Excel.Borders()..all=(Excel.Border()..lineStyle=Excel.LineStyle.thick);




 // String? path;
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text("Excel"),
        centerTitle: true,
      ),
      body: Center(
       child: Column(
         children: [
           ElevatedButton(
             onPressed: _createExcel,
             child: const Text("Excel"),
           ),
           const SizedBox(height: 40,),
           ElevatedButton(
             onPressed: _createPDF,
             child: const Text("PDf"),
           ),
         ],
       ),
      ),
    );
  }



  void _createExcel() async{
    Excel.Workbook workbook=Excel.Workbook();
    Excel.Worksheet sheet1=workbook.worksheets[0];


    List<Excel.ExcelDataRow> rows=List.generate(10, (index) =>
        Excel.ExcelDataRow(
            cells:<Excel.ExcelDataCell>[
              Excel.ExcelDataCell(columnHeader: 'Name', value:'name$index' ),
              Excel.ExcelDataCell(columnHeader: 'Date', value:'Date$index' ),
              Excel.ExcelDataCell(columnHeader: 'Address', value:'Address$index' ),
              Excel.ExcelDataCell(columnHeader: 'age', value:'1$index' ),
              Excel.ExcelDataCell(columnHeader: 'country', value:'country$index' ),
            ]
        )
    );



    sheet1.importData(rows, 2,2 );
    Excel.ExcelTable table= sheet1.tableCollection.create('customers', sheet1.getRangeByIndex(2, 2,2+rows.length,6));
    table.builtInTableStyle=Excel.ExcelTableBuiltInStyle.tableStyleLight5;


    List<int> bytes=workbook.saveAsStream();
    workbook.dispose();

    if(kIsWeb){
      AnchorElement(href: 'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
          ..setAttribute('download', 'document.xlsx')
          ..click();
    }else{
      String path=(await getApplicationSupportDirectory()).path;
      String fileName='$path/document.xlsx';
      File file=File(fileName);
      await file.writeAsBytes(bytes);
      OpenFile.open(fileName);
    }






  }



  void _createPDF() async{
    PdfDocument pdfDocument=PdfDocument();
    PdfPage page=pdfDocument.pages.add();



    page.graphics.drawString('Hello world', PdfStandardFont(PdfFontFamily.helvetica, 20));

    PdfGrid pdfGrid=PdfGrid();
    pdfGrid.style=PdfGridStyle(
      cellPadding: PdfPaddings(left: 4,top: 4,right: 4,bottom: 4),
      font: PdfStandardFont(PdfFontFamily.helvetica,25)
    );
    pdfGrid.columns.add(count: 4);
    pdfGrid.headers.add(1);

    PdfGridRow header=pdfGrid.headers[0];

    header.cells[0].value='Name';
    header.cells[1].value='Email';
    header.cells[2].value='age';
    header.cells[3].value='Address';
    PdfGridRow row;
    for(int i=0;i<10;i++){
     row=pdfGrid.rows.add();
      row.cells[0].value='name$i';
      row.cells[1].value='Email$i';
      row.cells[2].value='age$i';
      row.cells[3].value='Address$i';
    }

    pdfGrid.draw(page: page,bounds:Rect.fromLTWH(0, 50, 0, 0) );






    List<int> bytes=await pdfDocument.save();

    if(kIsWeb){
     AnchorElement(href:'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
     ..setAttribute('download','document.pdf' )
     ..click();
    }else{
      String path = (await getApplicationSupportDirectory()).path;
      String fileName = '$path/document.pdf';
      await File(fileName).writeAsBytes(bytes);
      OpenFile.open(fileName);
    }
  }
}
