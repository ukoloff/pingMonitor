//
// ����������� ���������� ����������� ������
//

var Hosts=[
{ip: '192.168.16.1', name: '����'},
{ip: '192.168.16.2', name: '���'},
{ip: '192.168.16.3', name: '���'},
{ip: 'e1.ru', name: 'e1'}
];

var $={};

//goW();
$.doc=newDoc();
startPage();

openLog();

while(!$.closed)
{
 pingOut();
 showHosts();
 WScript.Sleep(300);
}

closeLog()

//--[Functions]

// ��������������� � wscript (������ �������)
function goW()
{
 WScript.Interactive=false;

 if(WScript.FullName.replace(/^.*[\/\\]/, '').match(/^w/i)) return;
 (new ActiveXObject("WScript.Shell")).Run('wscript //B "'+
    WScript.ScriptFullName+'"', 0, false);
 WScript.Quit();
}

// ���������� ����������� � HTML-����
function html(S)
{
 S=''+S;
 var E={'&':'amp', '>':'gt', '<':'lt', '"':'quot'};
 for(var Z in E)
   S=S.split(Z).join('&'+E[Z]+';');
 return S;
}

// ������� MSIE
function newDoc()
{
 var ie=WScript.CreateObject('InternetExplorer.Application');
 ie.AddressBar=false;
 ie.StatusBar=false;
 ie.ToolBar=false;
 ie.MenuBar=false;
 ie.Visible=true;
 ie.Navigate('about:blank');
 while(ie.Busy) WScript.Sleep(100);
 $.ie=ie;
 return ie.Document;
}

// ������� ��������� ��������
function startPage()
{
 $.doc.open();
 $.doc.write(readSnippet('html'));
 $.doc.close();

 $.window=$.doc.parentWindow;
 $.doc.body.onunload=function(){ $.closed=1; }
 $.interior=$.doc.getElementById('Interior');
 insertHosts();
}

function insertHosts()
{
 for(var i in Hosts)
 {
  var H=Hosts[i];
  var r=$.interior.insertRow();
  r.insertCell().innerHTML='<BR />';
  r.insertCell().innerHTML=html(H.name);
  H.tr=r;
 }
}

// �������� ������� ������ �� ��������� ����
function readSnippet(name)
{
 var f=WScript.CreateObject("Scripting.FileSystemObject").
    OpenTextFile(WScript.ScriptFullName, 1);	//ForReading
 var on, R='';
 while(!f.AtEndOfStream)
 {
  var s=f.ReadLine();
  if(!on)
  {
   if(s.match(/^\s*\/\*[-\s]*\[([.\w]+)\][-\s]+$/i) && (RegExp.$1==name)) on=1;
   continue;
  }
  if(s.match(/^[-\s]+\*\/\s*$/)) break;
  R+=s+'\n';
 }
 f.Close();
 return R;
}

function openLog()
{
 var F=new ActiveXObject("Scripting.FileSystemObject");
 $.Log=F.OpenTextFile(F.GetParentFolderName(WScript.ScriptFullName)+'/'+
	F.GetBaseName(WScript.ScriptFullName)+'.log',
	8, /* ForAppending */
	true);
 writeLog('Started: '+WScript.ScriptFullName);
}

function writeLog(S)
{
 $.Log.WriteLine(new Date().N14()+'\t'+S);
}

function closeLog()
{
 writeLog('Stopped: '+WScript.ScriptFullName);
}

var sink_OnObjectReady, sink_OnCompleted, WMI, Sinc;

function pingOut()
{
 if(!WMI)WMI=GetObject("winmgmts:");
 if(!Sinc)Sinc=WScript.CreateObject("WbemScripting.SWbemSink", "sink_");

 if(!sink_OnObjectReady) sink_OnObjectReady=function(Ping, Ctx)
 {
  var H=Hosts[Ctx('i')];
  if((1!=H.stage) || H.wmi) return;
  H.wmi={
	code:	Ping.StatusCode,
	res:	Ping.PrimaryAddressResolutionStatus,
	ip:	Ping.ProtocolAddress,
	ms:	Ping.ResponseTime
  };
 }

 if(!sink_OnCompleted) sink_OnCompleted=function(hResult, lastError, Ctx)
 {
  Hosts[Ctx('i')].stage=2;
 }

 for(var i in Hosts)
 {
  var H=Hosts[i];
  if(1==H.stage) continue;
  if(!H.Ctx)
   (H.Ctx=WScript.CreateObject("WbemScripting.SWbemNamedValueSet")).
	Add('i', i);
  H.stage=1;
  H.wmi=0;
  WMI.ExecQueryAsync(Sinc,
	"Select * From Win32_PingStatus Where Timeout=300 "+
	"And Address='"+wmiEsc(H.ip)+"'",
	"WQL", 0, null, H.Ctx);
 } 
}

function wmiEsc(s)
{
 return (''+s).replace(/['\\]/g, '\\$&');
}

function showHosts()
{
 if($.closed) return;
 for(var i in Hosts)
 {
  var H=Hosts[i];
  if(!H.wmi) continue;
  var c=H.tr.cells[0];
  c.className=H.wmi.code? 'Red':'Green';
  if(!H.wmi.code) c.innerHTML=H.wmi.ms;
 }
}

function Number.prototype.N2()
{
 var N=''+this;
 while(N.length<2)N='0'+N;
 return N;
}

function Date.prototype.N14()
{
 return ''+this.getFullYear()+'-'+
    (this.getMonth()+1).N2()+'-'+
    this.getDate().N2()+'T'+
    this.getHours().N2()+':'+
    this.getMinutes().N2()+':'+
    this.getSeconds().N2();
}

//--[Snippets]

/*--[html]-----------------------------------------------------------
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<title>������� ����������� ����</title>

<style>

body	{
 background:	#A0C0E0;
 margin:	0;
 padding:	0.3ex;
 color:		black;
 font-family:	Verdana, Arial, sans-serif;
 text-align:	justify;
}

H1	{
 text-align:	right;
}

TD.Red	{
 background: red;
}

TD.Green {
 background: green;
}

</style>
</head><body>
<H1>������� ����������� ����</H1>

<Table id='Interior' Border Width='100%' CellSpacing='0'>
</Table>

</body></html>
-------------------------------------------------------------------*/

//--[EOF]------------------------------------------------------------