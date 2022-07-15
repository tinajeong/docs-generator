# Word 파일 POI로 수정하기

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
