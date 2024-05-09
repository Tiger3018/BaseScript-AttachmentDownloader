// @ts-nocheck
import { bitable } from '@lark-base-open/js-sdk';
//import { bitable } from '@base-open/web-api';  //旧版接口

import './index.scss';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
/////////获取html标签的定义区域///////////////////////////
//let oF = document.getElementById("oField");
let btn = document.getElementById("btn");
let xT = document.getElementById("xTable");
let xF = document.getElementById("xField");
let viewSelector=document.getElementById("view");
let qtyStatus = document.getElementById("qty");
let statusInfo=document.getElementById("info");
let mask = document.getElementById("mask");
let nameType=document.getElementById("nameType");
let downloadMethod=document.getElementById("downloadMethod");

//////////页面加载处理//////////////////
setLoading(true);

//获取当前多维表的数据表清单，以及各个数据表中包含有附件的字段内容
let oTables = await bitable.base.getTableMetaList();
oTables = oTables.filter(({name})=>name);
let allInfo=await getInfoByTableMetaList(oTables);
let allIndexFieldInfo=getIndexFieldMetaList(allInfo);

//检查字段中是否有附件类型字段
let allAttachmentFieldList=getAttachmentFieldMetaList(allInfo);

//根据是否有附件字段进行判断:
if (allAttachmentFieldList===false){ //当前多维表格无附件字段
  setLoading(false);
  alert("当前多维表格无附件类型字段");
}else{
  //填充数据表选择标记的选项
  fillSelect(xT, oTables);
  // 刚渲染本插件的时候，用户所选的tableId等信息
  const selection = await bitable.base.getSelection();
  //对表格字段选择控件，按照用户选择的数据表，设置初始的默认值
  xT.value = selection.tableId;
  
  //根据选择的数据表，更新字段选择控件的选项
  optionUpdateByTableSeleceted(xT.value,allAttachmentFieldList,"fieldMetaList",xF);
  //xF.value = xF.options['0'].value; //将第一个字段作为默认选项
  
  //根据选择的数据表，更新视图选择控件的选项
  optionUpdateByTableSeleceted(xT.value,allInfo,"viewMetaList",viewSelector);
  
  //为数据表选择select添加事件监听change，选择不同的数据表时，联动fields、视图的内容变更
  xT.addEventListener("change", optUpdate);

  //为按钮添加click监听
  btn.addEventListener("click", mainDownload);
  
  setLoading(false);
}
//页面加载完成，进入待用户操作过程/////////////

//////////////////////函数区域////////////////////////////////
//页面加载时的加载等待页面函数
function setLoading(loading) {
  mask.hidden = !loading;
}

//将获得的数组信息填充到select标记的option中的函数
function fillSelect(tMark, mArray) {
  for (let i = 0; i < mArray.length; i++) {
    let temp = document.createElement("option");
    temp.setAttribute("value", mArray[i].id);
    temp.textContent = mArray[i].name;
    tMark.appendChild(temp);
  }
}

//根据tableMetaList，输出对应的fieldMetaList
async function getInfoByTableMetaList(tableMetaList){
  let res=[];
  for (let i=0;i<tableMetaList.length;i++){
    let table=await bitable.base.getTableById(tableMetaList[i].id);
    let iFieldMetaList=await table.getFieldMetaList();
    let iViewMetaList=await table.getViewMetaList();
    let info={
      tableId: tableMetaList[i].id,
      tableName: tableMetaList[i].name,
      fieldMetaList: iFieldMetaList,
      viewMetaList: iViewMetaList
    };
    res.push(info);
  }
  return res;
}

//在输入的内容中，检查是否包含type===17的字段，如果包含，则返回一个对象数组。
//每个对象，包含tableId,tablName,fieldMetaList（只返回符合type==17的内容）
function getAttachmentFieldMetaList(info){
  let res=[];
  let hasAttachmentFieldMark=false;
  for (let i=0;i<info.length;i++){
    //从info中获取附件字段信息
    let iAttachmentFieldMetaList=info[i].fieldMetaList.filter(item=>item.type===17);
    //如果数组长度>0，说明有附件字段，则将标记置为true.
    if (iAttachmentFieldMetaList.length>0){
      hasAttachmentFieldMark=true;
    }
    //输出附件字段的MetaList信息。
    let attachmentFieldInfo={
      tableId: info[i].tableId,
      tableName: info[i].tableName,
      fieldMetaList:iAttachmentFieldMetaList
    };
    res.push(attachmentFieldInfo);
  }
  //循环结束后，判断附件字段标记，如果为true，则返回结果。如果为false，则返回fasle
  if (hasAttachmentFieldMark){
    return res;
  }else{
    return hasAttachmentFieldMark;
  }
}

//获得各个数据表的索引字段信息
function getIndexFieldMetaList(info){
  let res=[];
  for (let i=0;i<info.length;i++){
    let iIndexFieldMetaList=info[i].fieldMetaList.filter(item=>item.isPrimary===true);
    let indexFieldInfo={
      tableId: info[i].tableId,
      tableName: info[i].tableName,
      indexFieldId: iIndexFieldMetaList[0].id,
      indexFieldName: iIndexFieldMetaList[0].name
    };
    res.push(indexFieldInfo);
  }
  return res;
}

//根据数据表名，更新字段选择、视图选择的内容
/*
根据选择的数据表名称
从content中，筛选出该数据表名中的字段、视图MetaList
分别更新到fieldTag和viewTag的选项值。
*/
function optionUpdateByTableSeleceted(tableId,content,propName,tag){
  //待更新的option内容先置空。
  tag.options.length=0;
  let updatContent = content.filter(el=>el.tableId===tableId);
  if (updatContent.length !=0){
    for (let i=0;i<updatContent[0][propName].length;i++){
      let newopt = document.createElement("option");
      newopt.setAttribute("value", updatContent[0][propName][i].id);
      newopt.textContent = updatContent[0][propName][i].name;
      tag.appendChild(newopt);
    }
  }
}

//从数据表选择联动附件字段的函数，更新字段选项信息
function optUpdate() {
  //获取选择的数据表名称信息
  let selectedTableName=xT.value;
  if (selectedTableName!="0"){ //选中了某一个数据表
    //按照选中的数据表名称，更新字段和视图选择器。
    optionUpdateByTableSeleceted(selectedTableName,allAttachmentFieldList,"fieldMetaList",xF);
    optionUpdateByTableSeleceted(xT.value,allInfo,"viewMetaList",viewSelector);
  }
}

//点击按钮后的下载主函数。
async function mainDownload() {
  //获取用户选择的数据表Id
  let oTableId = xT.value;
  //获取用户选择的视图名称
  let oViewId=viewSelector.value;
  //获取用户选择的选项索引，0=原文件名，1=索引列名称
  let selectedNameType=nameType.selectedIndex;
  //获取用户选择的下载方式，0=逐个下载，1=zip下载
  let downloadMethodType=downloadMethod.selectedIndex;
  //获得需要操作的数据表
  let oTable = await bitable.base.getTableById(oTableId);
  //提取当前选中的oTable中的索引字段信息
  let indexFieldId=allIndexFieldInfo.filter(item=>item.tableId===oTableId)[0].indexFieldId;
  //获取用户选择的需要导出的字段Id
  let oFieldIdList = getMultiValue();//需要导出的字段是一个数组
  if (oFieldIdList.length==0){
    alert("未选择待下载字段");
  }
  let cellsInfo=await getCellsAddressList(oTable, oViewId, oFieldIdList); //获得待下载的单元格地址和总计文件数量。
  qtyStatus.textContent="共计有"+cellsInfo.fileQty+"个文件待下载。";

  //获得有附件的单元格行列地址信息
  let myCellsList=cellsInfo.cellsAddressList;
  //根据不同的选择模式，对单元格进行下载。
  //根据单元格地址数组，遍历获得待下载文件信息，并同步根据下载模式进行下载操作。
  //判断用户的操作选择方式
  let mode="";
  if (selectedNameType===0 && downloadMethodType===0){//使用原文件名,文件逐个下载
    mode="1";
  }else if (selectedNameType===0 && downloadMethodType===1){//使用原文件名,zip打包下载
    mode="2";
  }else if(selectedNameType===1 && downloadMethodType===0){//使用索引作为文件名,文件逐个下载
    mode="3";
  }else if(selectedNameType===1 && downloadMethodType===1){//使用索引作为文件名，zip打包下载
    mode="4";
  }

  //逐个下载的模式：
  if (mode==="1" || mode==="3"){
    for (let i=0;i<myCellsList.length;i++){
      let fileInfo=await getFileInfo(oTable,myCellsList[i].fieldId,myCellsList[i].recordId,indexFieldId);
      for (let ii=0;ii<fileInfo.length;ii++){
        if (mode==="1"){ //源文件名、逐个下载
          await download(fileInfo[ii]);
          //console.log("xxl", xxl);
          //statusInfo.textContent="完成第"+ii+"个";
        }else{//使用索引作为文件名,文件逐个下
          //将文件名替换成索引名称
          let nameInIndex=replaceFileName(fileInfo[ii].file_name,fileInfo[ii].index_info);
          fileInfo[ii].file_name=nameInIndex;
          await download(fileInfo[ii]);
        }
      } 
    }
  }


  //zip打包模式：
  let cnt=0;
  if (mode==="2" || mode==="4"){
    let jszip=new JSZip();
    let nameList=[];
    for (let i=0;i<myCellsList.length;i++){
      let fileInfo=await getFileInfo(oTable,myCellsList[i].fieldId,myCellsList[i].recordId,indexFieldId);
      
      for (let ii=0;ii<fileInfo.length;ii++){
        cnt=cnt+1;
        //将文件获得后打包到zip中
        let oUrl=fileInfo[ii].file_url;
        if (mode==="4"){
          //将文件名替换为索引的文件名
          let nameInIndex=replaceFileName(fileInfo[ii].file_name,fileInfo[ii].index_info);
          fileInfo[ii].file_name=nameInIndex;
        }
        //判断当前文件名是否与已有的冲突，如果冲突做更名处理。
        let checkedName=nameCheck(fileInfo[ii].file_name,fileInfo[ii].file_name,nameList,0);
        fileInfo[ii].file_name=checkedName;
        let data = await fetch(oUrl);
        let res=await data.blob();
        let blod = new Blob([res]);
        let oName = fileInfo[ii].file_name;
        jszip.file(oName,blod);
        //将已放入zip中的文件名存入数组，用于文件名重名判断使用。
        nameList.push(fileInfo[ii].file_name);
        statusInfo.textContent="加载第"+cnt+"个文件中……";
      }
    }
    //生成zip文件
    jszip.generateAsync({type:"blob"}, function updateCallback(metadata) {
      // 显示压缩进度
      var percent = metadata.percent.toFixed(2);
      statusInfo.textContent = "打包文件生成中："+percent + "%已完成。";
    })
      .then(function (zzfile) {
        saveAs(zzfile, "AttachmentPack.zip");
      });
  }
}

//获得所有待下载文件所在的单元格地址及文件数量信息。
//需要传入两个参数：数据表对象、[字段名称]数组。
async function getCellsAddressList(oTable, oViewId,fieldIdList){
  qtyStatus.textContent = "计算待下载文件数量……";
  let fileQty=0;
  let cellsAddressList=[];
  let oView=await oTable.getViewById(oViewId);
  //获取可见的记录条目数组
  let oRecordList = await oView.getVisibleRecordIdList();
  for (let i = 0; i < fieldIdList.length; i++) {
    //let oField = await oTable.getFieldById(fieldIdList[i]);
    //对oField字段中每一条Record进行循环，若不为空，则记录其单元格地址信息，并计数文件数量。
    for (let j = 0; j < oRecordList.length; j++) {
      let oCell = await oTable.getCellValue(fieldIdList[i], oRecordList[j]);
      if (oCell != null) {
        //将文件数量记录到计数器中
        fileQty = fileQty + oCell.length;
        //将cell的字段ID、记录ID保存下来，存入数组中。
        let cellAddress={fieldId:"",recordId:""};
        cellAddress.fieldId=fieldIdList[i];
        cellAddress.recordId=oRecordList[j];
        cellsAddressList.push(cellAddress);
        qtyStatus.textContent="计算待下载文件数量……"+fileQty+"个。若较慢，请刷新页面后重新尝试。";
      }
    }
  }
  return {cellsAddressList,fileQty};
}

/*
//根据输入的单元格行列信息，返回其中的文件信息(数组)
async function getFileInfo(oTable,fieldId,recordId,indexFieldId){
  let fileList=[];
  let oCell = await oTable.getCellValue(fieldId,recordId);
  let oCellIndexInfo= await oTable.getCellString(indexFieldId,recordId);
  for (let j=0;j<oCell.length;j++){
    let oToken = oCell[j].token;
    let oFileName = oCell[j].name;
    let oURL = await oTable.getAttachmentUrl(oToken, fieldId,recordId);
    //将获得的URL、文件名、索引字符串信息存放到一个文件对象中
    let fileInfo={file_url:"",file_name:"",index_info:""};
    fileInfo.file_url=oURL;
    fileInfo.file_name = oFileName;
    fileInfo.index_info=oCellIndexInfo;
    fileList.push(fileInfo);
  }
  return fileList;
}

*/

//使用getAttachmentUrls获取单元格内多附件的地址信息
async function getFileInfo(oTable,fieldId,recordId,indexFieldId){
  let fileList=[];
  let oCell = await oTable.getCellValue(fieldId,recordId);
  let oCellIndexInfo= await oTable.getCellString(indexFieldId,recordId);
  //将oCell中的所有token压入数组中
  let oTokens=[];
  for (let j=0;j<oCell.length;j++){
    oTokens.push(oCell[j].token);
  }
  let oUrls=await oTable.getCellAttachmentUrls(oTokens, fieldId, recordId);
  //将每一个附件的链接地址、文件信息放到fileInfo中，并压入fileList数组中
  for (let j=0;j<oCell.length;j++){
    let oFileName = oCell[j].name;
    let oURL =oUrls[j];
    //将获得的URL、文件名、索引字符串信息存放到一个文件对象中
    let fileInfo={file_url:"",file_name:"",index_info:""};
    fileInfo.file_url=oURL;
    fileInfo.file_name = oFileName;
    fileInfo.index_info=oCellIndexInfo;
    fileList.push(fileInfo);
  }
  return fileList;
}







//获取下拉多选内容的函数
function getMultiValue() {
  let multiV = [];
  let mseled = xF.selectedOptions;
  for (let i = 0; i < mseled.length; i++) {
    multiV.push(mseled[i].value);//取option的值
  }
  return multiV;
}


//检查文件名重复的函数，如果有重复的，添加顺序后缀。
function nameCheck(name,nextName,list,initialNum){
  let index=list.indexOf(nextName);
  if (index===-1){
    return nextName;
  }else if(index!=-1){
    //对名字添加后缀，再调用本函数比对
    initialNum=initialNum+1;
    let suffixStr="("+initialNum+")";
    let newName=changeFileName(name,suffixStr);
    let nn=nameCheck(name,newName,list,initialNum);
    return nn;
  }
}

//字符串内添加字符传信息
function insertStr (str, index, insertStr) {
   const ary = str.split('');		// 拆分，转化为数组
   ary.splice(index, 0, insertStr);	// 使用数组方法插入字符串
   return ary.join('');				// 拼接成字符串后输出
}

//在一个字符串中寻找某个标记所处的最后一个位置，若字符串中无mark标记，返回-1
function findLastMark(str,mark){
  let index = str.indexOf(mark);//首先找出第一次寻找到的位置
  let num =-1;
  while (index !== -1) {
    num=index;
    index = str.indexOf(mark, index + 1);//寻找下一个
  }
  return num;
}

//文件名加后缀更名函数
function changeFileName(nameStr,suffixStr){
  let temp=findLastMark(nameStr,".");
  let nStr="";
  //检查文件名字符串中是否包含.
  if (temp===-1){//文件名字符串无后缀，可直接添加
    nStr=nameStr+suffixStr;
  }else{//文件名有带点的信息，在最后一个点的位置添加
    nStr=insertStr(nameStr,temp,suffixStr);
  }
  return nStr;
}


//文件名替换函数
function replaceFileName(oNameStr,nNameStr){
  let temp=oNameStr.split(".");
  let nStr="";
  //检查文件名字符串中是否包含.
  if (temp.length===1){//源文件名中没有.，直接全量替换
    nStr=nNameStr;
  }else{//文件名有带点的信息，使用数组最后一个作为文件后缀，放弃其他的数组信息
    nStr=nNameStr+"."+temp[temp.length-1];
  }
  return nStr;
}


async function download(iFile) {
  let url=iFile.file_url;
  let filename=iFile.file_name;
  const response = await fetch(url);  
  if (!response.ok) {  
    throw new Error('Network response was not ok');  
  }else{
    statusInfo.textContent=filename+",文件生成中……";
    const blob = await response.blob();  
    const objectUrl = URL.createObjectURL(blob);     
    const a = document.createElement('a');  
    a.setAttribute('href', objectUrl);  
    a.setAttribute('download', filename);  
    a.click();
    //console.log("链接点击了",objectUrl);
    URL.revokeObjectURL(objectUrl);
    //console.log("链接取消了",objectUrl);
    return filename+"done";
  }
}


/*使用a标签下载
//若循环使用该方式，会遇到浏览器限制，需要设置延时。同时因为下载文件大小不确定，延时无法设置合理，不推荐使用。
function aDownload(iFile){
  let a = document.createElement('a');
  let fileName=iFile.file_name;
  let url=iFile.file_url;
  a.style.display = 'none';
  a.download = fileName;
  a.href = url;
  document.body.appendChild(a);
  a.click(); 
  document.body.removeChild(a);
}
*/