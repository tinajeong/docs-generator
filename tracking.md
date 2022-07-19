# Word 파일 POI로 생성/수정하기
- 테이블 분리해서 생성하기
- 샘플이미지 넣기
- 머릿글 추가하기
- 폰트 변경하기
- PDF로 변환하기
- 헤더 텍스트 수정하기

# 환경세팅
- https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=jwyoon25&logNo=221336857782
- https://www.tutorialspoint.com/apache_poi_word/apache_poi_word_quick_guide.htm


# 사용할 포맷
- XWPF (XML Word Processor Format) :  It is used to read and write **.docx** extension files of MS-Word.


# 의존성

```
Windows

Append the following strings to the end of the user variable

CLASSPATH −

C:\poi-bin-5.1.0\poi-5.1.0.jar;
C:\poi-bin-5.1.0\poi-ooxml-5.1.0.jar;
C:\poi-bin-5.1.0\poi-ooxml-full-5.1.0.jar;
C:\poi-bin-5.1.0\lib\commons-codec-1.15.jar;
C:\poi-bin-5.1.0\lib\commons-collections4-4.4.jar;
C:\poi-bin-5.1.0\lib\commons-io-2.11.0.jar;
C:\poi-bin-5.1.0\lib\commons-math3-3.6.1.jar;
C:\poi-bin-5.1.0\lib\log4j-api-2.14.1.jar;
C:\poi-bin-5.1.0\lib\SparseBitSet-1.2.jar;
C\poi-bin-5.1.0\ooxml-lib\commons-compress-1.21.jar
C\poi-bin-5.1.0\ooxml-lib\commons-logging-1.2.jar -- CVE 취약점
C\poi-bin-5.1.0\ooxml-lib\curvesapi-1.06.jar
C\poi-bin-5.1.0\ooxml-lib\slf4j-api-1.7.32.jar
C\poi-bin-5.1.0\ooxml-lib\xmlbeans-5.0.2.jar
```


# PDF 변환
- 의존성 호환이 안되는 이슈가 있는듯
> Exception in thread "main" java.lang.NoClassDefFoundError: com/lowagie/text/ExceptionConverter
```
org.apache.poi.xwpf.converter.pdf-1.0.6.jar
org.apache.poi.xwpf.converter.core-1.0.6.jar
ooxml-schemas-1.4.jar
itext-pdfa-5.5.13.jar
fr.opensagres.xdocreport.itext.extension-2.0.3.jar
```
