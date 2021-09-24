# java-excel-file-read-function part2

-기존에 poi라이브러리를 이용하여 엑셀 파일 읽기 프로그램에 대하여 설명하였다.

문제점: 엑셀 사이즈가 커지게 되면 엑셀 파일을 workbook 객체로 변환 할 때 Out of Memory(메모리 부족) 현상이 발생한다.<br>
해결 방안 : 대용량 데이터 처리 엑셀 양식을 이용 SAX 방식을 이용한다.

※ 사용되는 클래스

>SheetHandler.java(명칭은 커스텀한 것) -> 엑셀 양식에 맞게 데이터를 커스텀하여 정리하는 클래스

*SheetContentsHandler를 extends 받아 override를 하게 된다.
```C
  public class SheetHandler implements  XSSFSheetXMLMyHandler.SheetContentsHandler{
     
     private static final org.slf4j.Logger LOGGER =  LoggerFactory.getLogger(SheetHandler.class);
     //Row 에 해당한다
     private List<List<String>> rows = new ArrayList<List<String>>();
     //cell 값을 가지게 되는 배열
     private List<String> row = new ArrayList<String>(); 
    
     private List<String> header = new ArrayList<String>();

     //cell 위치
     private int currRowNum = 0;
     //공백 체크
     private int emptyCount = 10;

     //시작할때 위치를 알려준다.
     @Override
     public void startRow(int rowNum) {
           // int rowNum 은 첫 행을 읽을 때 위치를 말함 ex) 첫 행일 경우 0
           this.currRowNum = rowNum;
     }
          //cell의 끝자락 일 경우 발생하는 함수
          //첫 행에 해당하는 컬럼들의 값을 저장    
     @Override
     public void endRow(){
           this.rows.add(row); 
           row.clear();
     }
     
     //cell을 하나씩 읽어가는 함수
     @Override      //String formmatedValue 가 데이터에 해당 
     public void cell(String cellReference, String  formattedValue) {
           int thisCol = (new  CellReference(cellReference)).getCol(); // cell의 위치에 해당 ex) 첫번째 컬럼일 경우 0번에 해당    
           if("".equals(formattedValue)){
               row.add(""); 
            }else{
               row.add(formattedValue);
            }
     }
     
     @Override
     public void headerFooter(String text, boolean isHeader,  String tagName) {
           // TODO Auto-generated method stub
           //sheet의 첫 row 와 마지막 row를 처리하는 메소드
     }
}
```

> readLargeExcel(커스텀된 메소드로 SheetHandler 실행 하는 메소드이다)<br>

*SAX parsing 클래스(엑셀 파일을 읽어드리는 함수)

```C
//대용량 데이터 처리사용
     public static SheetHandler readLargeExcel(String  path,List<PublicCodeVO> publicCodelist) {
           SheetHandler sheetHandler = new  SheetHandler(publicCodelist);
           
                     // 현재 읽고자 하는 파일 가져오기
           File file = new File(path);
           try {
                //OPCPagkage 파일을 읽거나 쓸 수 있는 상태의 컨테이너를 생성
                OPCPackage opc = OPCPackage.open(file);
                //OPC 컨테이너를 XSSF형식으로 읽어 옴
                XSSFReader xssfReader = new XSSFReader(opc);
                //엑셀 스타일 형식을 가져오는건데.....
                StylesTable styles =  xssfReader.getStylesTable();
                
                ReadOnlySharedStringsTable strings = new  ReadOnlySharedStringsTable(opc);
                
                //엑셀의 시트를 하나만 가져옴
                //만양 여러 시트를 가져와야 할경우 while 문을 통해 처리
                //ex) xssfReader.getSheetsData().next(); 대신
                                // XSSFReader.Sheetlterator itr = (XSSFReader.Sheetlterator)xssfReader.getSheetsData(); -> sheet별로 collection으로 분할함
                                // while(itr.hasNext()) 를 통해 InputStream inputStream = itr.next(); 를 이용
                InputStream inputStream =  xssfReader.getSheetsData().next();
                
                InputSource inputSource = new  InputSource(inputStream);
                //직접적인 sheet의 cell과 row를 생성하는 이벤트
//              ContentHandler handle = new  XSSFSheetXMLHandler(styles, strings, sheetHandler, false);
                //XSSFSheetXMLHandler 를 재정의함
                ContentHandler handle = new  XSSFSheetXMLMyHandler(styles, strings, sheetHandler, false);
                
                SAXParserFactory saxFactory =  SAXParserFactory.newInstance();
                SAXParser saxParser = saxFactory.newSAXParser();
                
                XMLReader xmlReader = saxParser.getXMLReader();
                xmlReader.setContentHandler(handle);
                
                xmlReader.parse(inputSource);
                inputStream.close();
                opc.close();
                
           }catch (Exception e) {
                LOGGER.error("### excel read file error :  {}",e.getMessage());
                sheetHandler.setRows(null);
           }
           
           return sheetHandler;
     }
```

##SAX 파싱 방법 이용시 알아야 할점
>엑셀 양식이 어떻냐에 따라 다르겠지만 만약 한국식 날짜를 이용하는 경우 SAX 파싱을 이용할 경우 강제적으로 미국식 날짜를 가져오게된다. 
>> ex) 2021-01-01 -> 01/01/21로 강제 변환하여 들고 오게된다.
>> 이유 : SHeetContentsHandler는 **XSSFSheetXMLHandler**에 속하여 있다. **XSSFSheetXMLHandler**에서 실질적 데이터를 가져오게 된다.
>> XSSFSheetXMLHandler는 날짜 변환에 **DataFormatter**를 이용하고 있다. 
>> DataFormatter는 데이터에서 날짜가 어떤 locale을 가지던지 미국식 영어로 변환하여 가져오게 된다.
>> 만약 커스텀 마이징이 필요한 경우 DefaultHandler를 상속받아 클래스를 새로 정의 해줘 사용해야한다.

-> 커스텀 마이징이 되어야하는 
```C
   public void endElement(String uri, String localName,  String name)
                throws SAXException {
            String thisStr = null;
            // v => contents of a cell
            if (isTextTag(name)) {
                vIsOpen = false;
                
                // Process the value contents as required, now  we have it all
                switch (nextDataType) {
                    case BOOLEAN:
                        char first = value.charAt(0);
                        thisStr = first == '0' ? "FALSE" :  "TRUE";
                        break;
                    case ERROR:
                        thisStr = "ERROR:" + value.toString();
                        break;
                    case FORMULA:
                        if(formulasNotResults) {
                           thisStr = formula.toString();
                        } else {
                           String fv = value.toString();
                           
                           if (this.formatString != null) {
                              try {
                                 // Try to use the value as a  formattable number
                                 double d =  Double.parseDouble(fv);
                                 thisStr =  formatter.formatRawCellContents(d, this.formatIndex,  this.formatString);
                              } catch(NumberFormatException e)  {
                                 // Formula is a String result  not a Numeric one
                                 thisStr = fv;
                              }
                           } else {
                              // No formating applied, just do  raw value in all cases
                              thisStr = fv;
                           }
                        }
                        break;
                    case INLINE_STRING:
                        // TODO: Can these ever have formatting  on them?
                        XSSFRichTextString rtsi = new  XSSFRichTextString(value.toString());
                        thisStr = rtsi.toString();
                        break;
                    case SST_STRING:
                        String sstIndex = value.toString();
                        try {
                            int idx =  Integer.parseInt(sstIndex);
                            XSSFRichTextString rtss = new  XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
                            thisStr = rtss.toString();
                        }
                        catch (NumberFormatException ex) {
                            System.err.println("Failed to parse  SST index '" + sstIndex + "': " + ex.toString());
                        }
                        break;
                    case NUMBER:
                        String n = value.toString();
                        //formatIndex에 대하여 현재 설정된 번호에 한에서 한국식으로 받아 처리 할 수 있도록 커스터마이징을 진행 대부분 formatIndex 가 14일경우는 날짜로? 가져가는 거 같음
                        if (this.formatString != null) {    
                             //날짜 형식 관련 데이터들을  locale.Korea 형식으로 변경
                             if(formatIndex == 14 ||  formatIndex == 31 || formatIndex == 57 || formatIndex == 58 ||
                             (176 <= formatIndex && formatIndex  <=178) || (182 <= formatIndex && formatIndex <= 196) ||
                             (210 <= formatIndex && formatIndex  <=213) || (208 == formatIndex)) {
                                   sdf = new  SimpleDateFormat("yyyy-MM-dd");
                                   Date date =  DateUtil.getJavaDate(Double.parseDouble(n)); //해당값을 Date형으로 들고오게 함 그리하여 format을 통해 데이터를 재구성함
                                   thisStr = sdf.format(date);
                             }else {
                                   //Apache POI DataFormatter  formatter 는  로케일 형식 무시하고 미국 형식 날짜 표기법으로 반환함
                                   thisStr =  formatter.formatRawCellContents(Double.parseDouble(n),  this.formatIndex, this.formatString);
                             }
                        }else
                            thisStr = n;
                        break;
                    default:
                        thisStr = "(TODO: Unexpected type: " +  nextDataType + ")";
                        break;
                }
                
                // Output
                output.cell(cellRef, thisStr);
            } else if ("f".equals(name)) {
               fIsOpen = false;
            } else if ("is".equals(name)) {
               isIsOpen = false;
            } else if ("row".equals(name)) {
               output.endRow();
            }
            else if("oddHeader".equals(name) ||  "evenHeader".equals(name) ||
                  "firstHeader".equals(name)) {
               hfIsOpen = false;
               output.headerFooter(headerFooter.toString(),  true, name);
            }
            else if("oddFooter".equals(name) ||  "evenFooter".equals(name) ||
                  "firstFooter".equals(name)) {
               hfIsOpen = false;
               output.headerFooter(headerFooter.toString(),  false, name);
            }
        }

```
