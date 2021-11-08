[Ivy]
17CF0967FCDFADEC 9.3.0 #module
>Proto >Proto Collection #zClass
Rl0 ReadExcel Big #zClass
Rl0 B #cInfo
Rl0 #process
Rl0 @AnnotationInP-0n ai ai #zField
Rl0 @TextInP .type .type #zField
Rl0 @TextInP .processKind .processKind #zField
Rl0 @TextInP .xml .xml #zField
Rl0 @TextInP .responsibility .responsibility #zField
Rl0 @StartRequest f3 '' #zField
Rl0 @GridStep f0 '' #zField
Rl0 @PushWFArc f2 '' #zField
Rl0 @GridStep f1 '' #zField
Rl0 @PushWFArc f4 '' #zField
Rl0 @EndTask f5 '' #zField
Rl0 @PushWFArc f6 '' #zField
>Proto Rl0 Rl0 ReadExcel #zField
Rl0 f3 outLink start.ivp #txt
Rl0 f3 inParamDecl '<> param;' #txt
Rl0 f3 requestEnabled true #txt
Rl0 f3 triggerEnabled false #txt
Rl0 f3 callSignature start() #txt
Rl0 f3 caseData businessCase.attach=true #txt
Rl0 f3 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>ReadExcel.ivp</name>
    </language>
</elementInfo>
' #txt
Rl0 f3 @C|.responsibility Everybody #txt
Rl0 f3 73 49 30 30 -21 17 #rect
Rl0 f0 actionTable 'out=in;
out.filePath="excelFile.xls";
' #txt
Rl0 f0 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>init</name>
    </language>
</elementInfo>
' #txt
Rl0 f0 192 42 112 44 -8 -8 #rect
Rl0 f2 103 64 192 64 #arcP
Rl0 f1 actionTable 'out=in;
' #txt
Rl0 f1 actionCode 'import com.axonivy.connector.excel.ExcelReader;

out.recordset = new ExcelReader().getRecordset(in.filePath);' #txt
Rl0 f1 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>callReadExcel</name>
    </language>
</elementInfo>
' #txt
Rl0 f1 360 42 112 44 -39 -8 #rect
Rl0 f4 304 64 360 64 #arcP
Rl0 f5 529 49 30 30 0 15 #rect
Rl0 f6 472 64 529 64 #arcP
>Proto Rl0 .type com.axonivy.connector.excel.ReadExcelData #txt
>Proto Rl0 .processKind CALLABLE_SUB #txt
>Proto Rl0 0 0 32 24 18 0 #rect
>Proto Rl0 @|BIcon #fIcon
Rl0 f3 mainOut f2 tail #connect
Rl0 f2 head f0 mainIn #connect
Rl0 f0 mainOut f4 tail #connect
Rl0 f4 head f1 mainIn #connect
Rl0 f1 mainOut f6 tail #connect
Rl0 f6 head f5 mainIn #connect
