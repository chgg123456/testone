function 获取文件夹下文件路径()
{
Application.ScreenUpdating=false;
var fb=Application.FileDialog(msoFileDialogFolderPicker);
fb.Show();
var p=fb.SelectedItems(1);
Application.ScreenUpdating=true; 
if(p==null) return; 
//Range('a:a').Clear(); 
var f=Dir(p+'\\');//遍历文件夹中的所有文件
let arr=[]
var i=1;
while(f){ //按文件名数量循环

//ActiveSheet.Cells(i,1).Value2=p+'\\'+f;
if(f.endsWith('.docx')){arr.push(f)}
i++; //可以求文件夹下文件个数

f=Dir() //获取下一个文件名
//console.log(f);
}
arr.sort((a,b)=>a.slice(18,20)-b.slice(18,20))
arr=arr.map(item=>p+'\\'+item)
//console.log(JSON.stringify(arr))
return arr
//Documents.Open(p+'\\'+arr[0])
}
