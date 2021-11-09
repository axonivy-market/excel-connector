[Ivy]
17CFF1C40579494A 9.3.0 #module
>Proto >Proto Collection #zClass
Ps0 PersonEditorProcess Big #zClass
Ps0 RD #cInfo
Ps0 #process
Ps0 @AnnotationInP-0n ai ai #zField
Ps0 @TextInP .type .type #zField
Ps0 @TextInP .processKind .processKind #zField
Ps0 @TextInP .xml .xml #zField
Ps0 @TextInP .responsibility .responsibility #zField
Ps0 @UdInit f0 '' #zField
Ps0 @UdProcessEnd f1 '' #zField
Ps0 @UdEvent f3 '' #zField
Ps0 @UdExitEnd f4 '' #zField
Ps0 @PushWFArc f5 '' #zField
Ps0 @GridStep f6 '' #zField
Ps0 @PushWFArc f7 '' #zField
Ps0 @PushWFArc f2 '' #zField
Ps0 @UdEvent f8 '' #zField
Ps0 @UdProcessEnd f9 '' #zField
Ps0 @GridStep f11 '' #zField
Ps0 @PushWFArc f12 '' #zField
Ps0 @CallSub f13 '' #zField
Ps0 @PushWFArc f14 '' #zField
Ps0 @UdEvent f18 '' #zField
Ps0 @UdProcessEnd f19 '' #zField
Ps0 @GridStep f21 '' #zField
Ps0 @PushWFArc f22 '' #zField
Ps0 @PushWFArc f10 '' #zField
Ps0 @UdEvent f26 '' #zField
Ps0 @UdProcessEnd f27 '' #zField
Ps0 @CallSub f29 '' #zField
Ps0 @PushWFArc f30 '' #zField
Ps0 @GridStep f31 '' #zField
Ps0 @PushWFArc f32 '' #zField
Ps0 @PushWFArc f25 '' #zField
Ps0 @GridStep f15 '' #zField
Ps0 @PushWFArc f16 '' #zField
Ps0 @PushWFArc f17 '' #zField
>Proto Ps0 Ps0 PersonEditorProcess #zField
Ps0 f0 guid 17CFF1C405BEC8A7 #txt
Ps0 f0 method start() #txt
Ps0 f0 inParameterDecl '<> param;' #txt
Ps0 f0 outParameterDecl '<> result;' #txt
Ps0 f0 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>start()</name>
    </language>
</elementInfo>
' #txt
Ps0 f0 83 51 26 26 -16 15 #rect
Ps0 f1 339 51 26 26 0 12 #rect
Ps0 f3 guid 17CFF1C406113515 #txt
Ps0 f3 actionTable 'out=in;
' #txt
Ps0 f3 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>close</name>
    </language>
</elementInfo>
' #txt
Ps0 f3 83 147 26 26 -15 15 #rect
Ps0 f4 211 147 26 26 0 12 #rect
Ps0 f5 109 160 211 160 #arcP
Ps0 f6 actionTable 'out=in;
' #txt
Ps0 f6 actionCode 'import com.axonivy.connector.excel.demo.Person;

Person p = new Person();
p.firstname = "Louis";
p.lastname = "Müller";
out.persons.add(p);

p = new Person();
p.firstname = "Reguel";
p.lastname = "Wermelinger";
out.persons.add(p);

p = new Person();
p.firstname = "Alexander";
p.lastname = "Suter";
out.persons.add(p);' #txt
Ps0 f6 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>init persons</name>
    </language>
</elementInfo>
' #txt
Ps0 f6 168 42 112 44 -32 -8 #rect
Ps0 f7 109 64 168 64 #arcP
Ps0 f2 280 64 339 64 #arcP
Ps0 f8 guid 17CFF243C057A114 #txt
Ps0 f8 actionTable 'out=in;
' #txt
Ps0 f8 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>export</name>
    </language>
</elementInfo>
' #txt
Ps0 f8 83 211 26 26 -16 15 #rect
Ps0 f9 691 211 26 26 0 12 #rect
Ps0 f11 actionTable 'out=in;
' #txt
Ps0 f11 actionCode 'import com.axonivy.connector.excel.demo.Person;

out.exportRecordset = new Recordset();

for (Person p : in.persons){
 Record r = new Record(["firstname","lastname"], [p.firstname,p.lastname]);
 out.exportRecordset.add(r);
}' #txt
Ps0 f11 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>recordset mapping</name>
    </language>
</elementInfo>
' #txt
Ps0 f11 168 202 112 44 -52 -8 #rect
Ps0 f12 109 224 168 224 #arcP
Ps0 f13 processCall excel_connector/WriteExcel:write(Recordset) #txt
Ps0 f13 requestActionDecl '<Recordset recordset> param;' #txt
Ps0 f13 requestMappingAction 'param.recordset=in.exportRecordset;
' #txt
Ps0 f13 responseMappingAction 'out=in;
out.excelFile=result.excelFile;
' #txt
Ps0 f13 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>excel_connector/WriteExcel</name>
    </language>
</elementInfo>
' #txt
Ps0 f13 304 202 160 44 -74 -8 #rect
Ps0 f14 280 224 304 224 #arcP
Ps0 f18 guid 17CFF85A10B9B06F #txt
Ps0 f18 actionTable 'out=in;
' #txt
Ps0 f18 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>addPerson</name>
    </language>
</elementInfo>
' #txt
Ps0 f18 83 315 26 26 -30 15 #rect
Ps0 f19 339 315 26 26 0 12 #rect
Ps0 f21 actionTable 'out=in;
' #txt
Ps0 f21 actionCode 'import java.io.FileInputStream;
import java.io.InputStream;
import org.primefaces.model.DefaultStreamedContent;

InputStream stream;
stream = new FileInputStream(in.excelFile.getJavaFile());
out.streamedExcelFile = new DefaultStreamedContent(stream,"application/vnd.ms-excel","demo_export.xls");' #txt
Ps0 f21 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>set streamed file</name>
    </language>
</elementInfo>
' #txt
Ps0 f21 520 202 112 44 -46 -8 #rect
Ps0 f22 464 224 520 224 #arcP
Ps0 f10 632 224 691 224 #arcP
Ps0 f26 guid 17D03B6D64213F9C #txt
Ps0 f26 actionTable 'out=in;
' #txt
Ps0 f26 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>importExcel</name>
    </language>
</elementInfo>
' #txt
Ps0 f26 75 403 26 26 -32 15 #rect
Ps0 f27 547 403 26 26 0 12 #rect
Ps0 f29 processCall excel_connector/ReadExcel:read(File) #txt
Ps0 f29 requestActionDecl '<File excelFile> param;' #txt
Ps0 f29 requestMappingAction 'param.excelFile=in.importFile;
' #txt
Ps0 f29 responseMappingAction 'out=in;
out.importRecordset=result.recordset;
' #txt
Ps0 f29 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>excel_connector/ReadExcel</name>
    </language>
</elementInfo>
' #txt
Ps0 f29 144 394 160 44 -75 -8 #rect
Ps0 f30 304 416 360 416 #arcP
Ps0 f31 actionTable 'out=in;
' #txt
Ps0 f31 actionCode 'import com.axonivy.connector.excel.demo.Person;

for(Record r : in.importRecordset.getRecords()){
	Person p = new Person();
	p.firstname = r.getField("firstname").toString();
	p.lastname = r.getField("lastname").toString();
	out.persons.add(p);
}' #txt
Ps0 f31 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>add records to persons</name>
    </language>
</elementInfo>
' #txt
Ps0 f31 360 394 144 44 -64 -8 #rect
Ps0 f32 504 416 547 416 #arcP
Ps0 f32 0 0.24154701666138945 0 0 #arcLabel
Ps0 f25 101 416 144 416 #arcP
Ps0 f15 actionTable 'out=in;
' #txt
Ps0 f15 actionCode 'import com.axonivy.connector.excel.demo.Person;

Person p = new Person();
p.firstname = in.firstname;
p.lastname = in.lastname;
out.persons.add(p);

in.firstname = "";
in.lastname = "";' #txt
Ps0 f15 @C|.xml '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<elementInfo>
    <language>
        <name>append person</name>
    </language>
</elementInfo>
' #txt
Ps0 f15 168 306 112 44 -42 -8 #rect
Ps0 f16 109 328 168 328 #arcP
Ps0 f17 280 328 339 328 #arcP
>Proto Ps0 .type com.axonivy.connector.excel.demo.PersonEditor.PersonEditorData #txt
>Proto Ps0 .processKind HTML_DIALOG #txt
>Proto Ps0 -8 -8 16 16 16 26 #rect
Ps0 f3 mainOut f5 tail #connect
Ps0 f5 head f4 mainIn #connect
Ps0 f0 mainOut f7 tail #connect
Ps0 f7 head f6 mainIn #connect
Ps0 f6 mainOut f2 tail #connect
Ps0 f2 head f1 mainIn #connect
Ps0 f8 mainOut f12 tail #connect
Ps0 f12 head f11 mainIn #connect
Ps0 f11 mainOut f14 tail #connect
Ps0 f14 head f13 mainIn #connect
Ps0 f13 mainOut f22 tail #connect
Ps0 f22 head f21 mainIn #connect
Ps0 f21 mainOut f10 tail #connect
Ps0 f10 head f9 mainIn #connect
Ps0 f29 mainOut f30 tail #connect
Ps0 f30 head f31 mainIn #connect
Ps0 f31 mainOut f32 tail #connect
Ps0 f32 head f27 mainIn #connect
Ps0 f26 mainOut f25 tail #connect
Ps0 f25 head f29 mainIn #connect
Ps0 f18 mainOut f16 tail #connect
Ps0 f16 head f15 mainIn #connect
Ps0 f15 mainOut f17 tail #connect
Ps0 f17 head f19 mainIn #connect
