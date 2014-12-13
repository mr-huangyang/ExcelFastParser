package com.oy.excel.fast.parser;

import com.oy.excel.fast.parser.row.XRow;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2014-12-7.
 */
public class Excel972003Parser implements HSSFListener {

    private SSTRecord sstrec;
    private int curRealRowIndex ;
    private int previewRowIndex;
    private int curSheet = 0;
    private Map<String,List<XRow>> sheetRows ;
    public Excel972003Parser(){
        curRealRowIndex = 0;
        previewRowIndex = 0;
        sheetRows = new HashMap<String, List<XRow>>();
    }
    @Override
    public void processRecord(Record record) {
              XRow xr;
              switch (record.getSid()){
                  case BOFRecord.sid:
                      BOFRecord bof = (BOFRecord) record;
                      if (bof.getType() == BOFRecord.TYPE_WORKBOOK) {
                          System.out.println("Encountered workbook");
                      } else if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {
                          curSheet++;
                          sheetRows.put(curSheet+"",new ArrayList<XRow>());
                      }
                      break;
                  case BoundSheetRecord.sid :
                      BoundSheetRecord bsr = (BoundSheetRecord) record;
                      break;
                  case SSTRecord.sid:
                      sstrec = (SSTRecord) record;
                      break;
                  case RowRecord.sid: //row has real columns
                      RowRecord rowRecord = (RowRecord)record;
                      this.curRealRowIndex = rowRecord.getRowNumber();
                      //handle the dummy row
                      if(this.previewRowIndex + 1 < this.curRealRowIndex ){
                          for(int i = this.previewRowIndex ; i < this.curRealRowIndex ;i++ ){
                              xr = new XRow(i+1);
                              xr.setNotExist(true);
                              this.sheetRows.get(curSheet+"").add(xr);
                          }
                      }else{
                          xr = new XRow(this.curRealRowIndex);
                          this.sheetRows.get(curSheet+"").add(xr);
                      }
                       this.previewRowIndex = this.curRealRowIndex;
                      break;

                  case LabelSSTRecord.sid: // not blank record
                      LabelSSTRecord lrec = (LabelSSTRecord) record;
                      int rowIndex =   lrec.getRow();
                      XRow curXRow = this.sheetRows.get(curSheet+"").get(rowIndex);
                      curXRow.setCellValue(sstrec.getString(lrec.getSSTIndex()).toString());
                      break;
                  case NumberRecord.sid:
                      NumberRecord numberRecord = (NumberRecord) record;
                      this.sheetRows.get(curSheet+"").get(numberRecord.getRow()).setCellValue(numberRecord.getValue()+"");
                      break;
                  default:
                      break;
              }


    }

    //@Override
    public void processRecord1(Record record) {
        switch (record.getSid()) {
            // the BOFRecord can represent either the beginning of a sheet or the workbook
            case BOFRecord.sid:
                BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == BOFRecord.TYPE_WORKBOOK) {
                    System.out.println("Encountered workbook");
                    // assigned to the class level member
                } else if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {
                    System.out.println("Encountered sheet reference");
                }
                break;
            case BoundSheetRecord.sid:
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                System.out.println("New sheet named: " + bsr.getSheetname());

                break;
            case RowRecord.sid:
                RowRecord rowrec = (RowRecord) record;
                System.out.println(rowrec.getRowNumber()+1 +
                        "Row found, first column at "
                        + rowrec.getFirstCol() + " last column at " + rowrec.getLastCol());
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;
                System.out.println("Cell found with value " + numrec.getValue()
                        + " at row " + numrec.getRow() + " and column " + numrec.getColumn());
                break;
            // SSTRecords store a array of unique strings used in Excel.
            case SSTRecord.sid:
                sstrec = (SSTRecord) record;
                for (int k = 0; k < sstrec.getNumUniqueStrings(); k++) {
                    System.out.println("String table value " + k + " = " + sstrec.getString(k));
                }
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lrec = (LabelSSTRecord) record;
                System.out.println("String cell found with value "
                        + sstrec.getString(lrec.getSSTIndex()));
                break;
        }
    }

    public static void main(String[] args) throws Exception {
        // create a new file input stream with the input file specified
        // at the command line
        FileInputStream fin = new FileInputStream("c:\\Users\\Administrator\\Desktop\\bb.xls");
        // create a new org.apache.poi.poifs.filesystem.Filesystem
        POIFSFileSystem poifs = new POIFSFileSystem(fin);
        // get the Workbook (excel part) stream in a InputStream
        InputStream din = poifs.createDocumentInputStream("Workbook");
        // construct out HSSFRequest object
        HSSFRequest req = new HSSFRequest();
        // lazy listen for ALL records with the listener shown above
        Excel972003Parser parser = new Excel972003Parser();
        req.addListenerForAllRecords(parser);
        // create our event factory
        HSSFEventFactory factory = new HSSFEventFactory();
        // process our events based on the document input stream
        factory.processEvents(req, din);
        // once all the events are processed close our file input stream
        fin.close();
        // and our document input stream (don't want to leak these!)
        din.close();
        System.out.println("done.");
        List<XRow> rows1 = parser.sheetRows.get(1+"");
        for(XRow x : rows1){
            x.toString();
        }
        rows1 = parser.sheetRows.get(2+"");
//        for(XRow x : rows1){
//            x.toString();
//        }
    }
}
