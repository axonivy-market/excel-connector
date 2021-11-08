[Ivy]
17CF0992EAD1CDCB 9.3.0 #module
>Proto >Proto Collection #zClass
Wl0 WriteExcel Big #zClass
Wl0 B #cInfo
Wl0 #process
Wl0 @AnnotationInP-0n ai ai #zField
Wl0 @TextInP .type .type #zField
Wl0 @TextInP .processKind .processKind #zField
Wl0 @TextInP .xml .xml #zField
Wl0 @TextInP .responsibility .responsibility #zField
Wl0 @StartRequest f3 '' #zField
Wl0 @GridStep f0 '' #zField
Wl0 @PushWFArc f1 '' #zField
Wl0 @GridStep f2 '' #zField
Wl0 @EndTask f4 '' #zField
Wl0 @PushWFArc f5 '' #zField
Wl0 @PushWFArc f6 '' #zField
>Proto Wl0 Wl0 WriteExcel #zField
Wl0 f3 outLink start.ivp #txt
Wl0 f3 inParamDecl '<> param;' #txt
Wl0 f3 requestEnabled true #txt
Wl0 f3 triggerEnabled false #txt
Wl0 f3 callSignature start() #txt
Wl0 f3 caseData businessCase.attach=true #txt
Wl0 f3 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>writeExcel.ivp</name>
    </language>
</elementInfo>
' #txt
Wl0 f3 @C|.responsibility Everybody #txt
Wl0 f3 57 65 30 30 -21 17 #rect
Wl0 f0 actionTable 'out=in;
out.filePath="excelFile.xls";
' #txt
Wl0 f0 actionCode 'out.recordset = new Recordset();
out.recordset.addColumn("a1",["ad","sg"]);
out.recordset.addColumn("a2",[45,1234]);
out.recordset.addColumn("a3",[new Date(), new DateTime()]);' #txt
Wl0 f0 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>init</name>
    </language>
</elementInfo>
' #txt
Wl0 f0 136 58 112 44 -8 -8 #rect
Wl0 f1 87 80 136 80 #arcP
Wl0 f2 actionTable 'out=in;
' #txt
Wl0 f2 actionCode 'import com.axonivy.connector.excel.ExcelWriter;

new ExcelWriter().execute(in.recordset, in.filePath);' #txt
Wl0 f2 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>call WriteExcel</name>
    </language>
</elementInfo>
' #txt
Wl0 f2 312 58 112 44 -39 -8 #rect
Wl0 f4 497 65 30 30 0 15 #rect
Wl0 f5 248 80 312 80 #arcP
Wl0 f6 424 80 497 80 #arcP
>Proto Wl0 .type com.axonivy.connector.excel.WriteExcelData #txt
>Proto Wl0 .processKind CALLABLE_SUB #txt
>Proto Wl0 0 0 32 24 18 0 #rect
>Proto Wl0 @|BIcon #fIcon
Wl0 f3 mainOut f1 tail #connect
Wl0 f1 head f0 mainIn #connect
Wl0 f0 mainOut f5 tail #connect
Wl0 f5 head f2 mainIn #connect
Wl0 f2 mainOut f6 tail #connect
Wl0 f6 head f4 mainIn #connect
