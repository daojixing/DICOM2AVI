''''''''''参数修改区域''''''''''''''''''''''''''''''''
questions="A"    '此处B  是问题所在列
options="B"      '此处C 是选项所在列
num=341           '题库题数
kejian=false    'word  窗口是否可见
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''下面示答案设置项''''''''''''''''''''''''''''''''''''''''''''
answers="C"       '此处填写答案所在列  没有答案则留空
deleteas=true   '此处true 则删除干扰项，此处false保留干扰项
xiechuas=false     '此处true  则写出答案，false不写出答案。
blod=false          '答案项是否加粗，true为加粗，false 不加粗
highlight=7       '答案项背景颜色  7为黄色背景，8无背景
mark="|"         '选项分割
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''正则配置区'''''''''''''''''''''''''''''''',,
patrn="<img.*?>"   '正则表达式  正则表达式  
replStr=""          '替换成
''''''''''''''''''''配置区'''''''''''''''''''''''''''''''


SET Wshell=CreateObject("Wscript.Shell")
  if Wscript.Arguments.Count =0 then 
    msgbox ""&vblf&"请把excel文件拖到此程序上"
    WScript.Quit
  end if

Sub forceCScriptExecution
    Dim Arg, Str
    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
            If InStr( Arg, " " ) Then Arg = """" & Arg & """"
            Str = Str & " " & Arg
        Next
        CreateObject( "WScript.Shell" ).Run _
            "cscript //nologo """ & _
            WScript.ScriptFullName & _
            """ " & Str
        WScript.Quit
    End If
End Sub
forceCScriptExecution
Wscript.echo "·········配置如下················"
Wscript.echo "问题列:  "&questions
Wscript.echo "选项列： "&options
Wscript.echo "答案列： "&answers
Wscript.echo "题数：   "&num
Wscript.echo "写出答案：  "&xiechuas
Wscript.echo "删除干扰项："&deleteas
Wscript.echo "word窗口可见："&kejian
Wscript.echo "答案项加粗：  "&blod
Wscript.echo "答案项高亮：  "&highlight&"   注：7为黄色，8取消高亮"
Wscript.echo "·································"
For i=0 To WScript.Arguments.Count-1
filepath=WScript.Arguments(i)
Wscript.echo filepath
next
bool=MsgBox("是否继续",vbYesNo)
If bool=6 Then
Else
Wscript.echo "运行结束"
WScript.Quit
End If

For i=0 To WScript.Arguments.Count-1
filepath=WScript.Arguments(i) 
Wscript.echo "创建word对象"
Set wdObj = CreateObject("Word.Application")
Wscript.echo "创建excel对象"
set oExcel=CreateObject("Excel.Application")
main(filepath)
next

sub main(filepath)
wdObj.Visible=kejian
set objdoc=wdObj.Documents.Add
Wscript.echo "创建word文档"
set fso=createobject("scripting.filesystemobject")
Set oWb=oExcel.Workbooks.Open(filepath)
set oSheet=oWb.Sheets("Sheet1")
Wscript.echo "读取excel"
for x=1 to num
Wscript.echo "读取第"&x&"题"
geshi_cancel

wdObj.selection.TypeText x&": "&oSheet.Range(questions&x).Value&vbLf   '写出问题
str=split(oSheet.Range(options&x).Value,"|")                            '分隔答案项
If answers="" Then                                                     '如果没有定义答案列
asnull(str)                 
Else
  asr=oSheet.Range(answers&x).Value                     '答案
  If xiechuas Then
  wdObj.selection.TypeText   asr&vbLf     '写出答案
  End if
  If deleteas Then
  delteas str,asr
  Else
  asw str,asr 
  End If
End If 
next
oExcel.Workbooks.Close
oExcel.Quit
Wscript.echo "关闭excel"
set oExcel=Nothing
ql()

docfile=Left(filepath, InStrRev(filepath,".")-1)&".doc"
Wscript.echo "word保存在"&docfile 
objdoc.SaveAs(docfile)
objdoc.Close
Wscript.echo "关闭word"
wdObj.Quit
Set wdObj=Nothing
Wscript.echo "运行结束请关闭黑框"
end Sub

Sub asnull(str)

for i=0 to ubound(str)
Select Case i
	case 0  wdObj.selection.TypeText "A "&str(i)&vbLf
	case 1 wdObj.selection.TypeText "B "&str(i)&vbLf
	case 2 wdObj.selection.TypeText "C "&str(i)&vbLf
	case 3 wdObj.selection.TypeText "D "&str(i)&vbLf
	case 4 wdObj.selection.TypeText "E "&str(i)&vbLf
End Select
next
End Sub

Sub delteas (str,asr)              '写出答案删除干扰项
for i=0 to ubound(str)
Select Case i
	case 0  
	if ubound(split(asr,"A"))>0 then
	
	
	geshi
	wdObj.selection.TypeText "A "&str(i)&vbLf
	end if
	case 1 
	if ubound(split(asr,"B"))>0 then
	geshi
	wdObj.selection.TypeText "B "&str(i)&vbLf
	end if
	case 2 
	if ubound(split(asr,"C"))>0 then
	geshi
	wdObj.selection.TypeText "C "&str(i)&vbLf
	end if
	case 3 
	if ubound(split(asr,"D"))>0 then
	geshi
	wdObj.selection.TypeText "D "&str(i)&vbLf
	end if
	case 4
	if ubound(split(asr,"E"))>0 then
	geshi
   	wdObj.selection.TypeText "E "&str(i)&vbLf
	end if
End Select
next
End Sub

Sub asw (str,asr)                              '写出答案
for i=0 to ubound(str)
  Select Case i
	case 0  
	if ubound(split(asr,"A"))>0 then
	geshi
	else
	geshi_cancel
	end if
	wdObj.selection.TypeText "A "&str(i)&vbLf
	case 1 
	if ubound(split(asr,"B"))>0 then
	geshi
	else
	geshi_cancel
	end if
	wdObj.selection.TypeText "B "&str(i)&vbLf
	case 2 
	if ubound(split(asr,"C"))>0 then
 	geshi
	else
	geshi_cancel
	end if
	wdObj.selection.TypeText "C "&str(i)&vbLf
	case 3 
	if ubound(split(asr,"D"))>0 then
	geshi
	else
	geshi_cancel
	end if
	wdObj.selection.TypeText "D "&str(i)&vbLf
	case 4
	if ubound(split(asr,"E"))>0 then
	'Selection.Font.Underline =true 
	geshi
	else
	'Selection.Font.Underline =false
	geshi_cancel
	end if
	wdObj.selection.TypeText "E "&str(i)&vbLf
  End Select
next

End Sub


Sub geshi   '格式
wdObj.Selection.range.highlightcolorindex=highlight
wdObj.Selection.Font.Bold=blod
End Sub
Sub geshi_cancel
wdObj.Selection.range.highlightcolorindex=8
wdObj.Selection.Font.Bold=false
End Sub



Sub ql()
Wscript.echo "正在替换img字符串。。。。。。。"
wdObj.Selection.WholeStory
str2=wdObj.Selection.text
Set regEx = New RegExp              
regEx.Pattern = patrn              
regEx.IgnoreCase = True   
regEx.global=true
set matchs=regEx.execute(str2)
for each match in matchs
Wscript.echo "替换"&match
match=Replace(match," />","")
match=Replace(match,"<","")
With wdObj.selection.Find
.Text=match
.Replacement.Text=""
.Forward=True
.Wrap=wdFindContinue
.MatchWildcards=True
.Execute ,,,,,,,,,,2
End With
next 
With wdObj.selection.Find
.Text="< />"
.Replacement.Text=""
.Forward=True
.Wrap=wdFindContinue
.MatchWildcards=True
.Execute ,,,,,,,,,,2
End With
Wscript.echo "替换（）"
With wdObj.selection.Find
.Text="\(    \)"
.Replacement.Text=""
.Forward=True
.Wrap=wdFindContinue
.MatchWildcards=True
.Execute ,,,,,,,,,,2
End With
Wscript.echo "替换<>"
With  wdObj.selection.Find
.Text="\< /\>"
.Replacement.Text=""
.Forward=True
.Wrap=wdFindContinue
.MatchWildcards=True
.Execute ,,,,,,,,,,2
End With

for i=1 to 10
Wscript.echo "替换上下标"&i
Subscript(i)
Supscript(i)
next
Supscript("-")
'<sup>-</sup>
End Sub

sub Subscript(number)
With  wdObj.selection.Find
.Text="\<sub\>"&number&"\</sub\>"
.Replacement.Text=number
.Replacement.font.Subscript=true
.Forward=True
.Wrap=wdFindContinue
.MatchWildcards=True
.Execute ,,,,,,,,,,2
End With
end sub

sub Supscript(number)
With  wdObj.selection.Find
.Text="\<sup\>"&number&"\</sup\>"
.Replacement.Text=number
.Replacement.font.Superscript=true
.Forward=True
.Wrap=wdFindContinue
.MatchWildcards=True
.Execute ,,,,,,,,,,2
End With
end sub

