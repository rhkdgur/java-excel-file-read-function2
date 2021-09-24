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

*SAX parsing 클래스(엑셀 파일을 읽어드리는 함수)
