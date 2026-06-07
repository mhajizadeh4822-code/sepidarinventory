
let masterItems=[];
let counts=JSON.parse(localStorage.getItem('counts')||'[]');

excelInput.onchange=e=>{
 const r=new FileReader();
 r.onload=()=>{
  const wb=XLSX.read(r.result,{type:'binary'});
  const data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  masterItems=data;
  alert('کالاها بارگذاری شد');
 };
 r.readAsBinaryString(e.target.files[0]);
};

searchInput.addEventListener('input',()=>{
 let q=searchInput.value.toLowerCase();
 searchResults.innerHTML='';
 masterItems.filter(x=>String(x.ItemName||'').toLowerCase().includes(q)||String(x.ItemCode||'').includes(q))
 .slice(0,10).forEach(i=>{
   let d=document.createElement('div');
   d.innerText=i.ItemCode+' - '+i.ItemName;
   d.onclick=()=>{
    itemCode.value=i.ItemCode;
    itemName.value=i.ItemName;
   };
   searchResults.appendChild(d);
 });
});

function saveCount(){
 if(!itemCode.value || Number(qty.value)<=0){alert('اطلاعات ناقص است');return;}
 counts.push({code:itemCode.value,name:itemName.value,qty:qty.value});
 localStorage.setItem('counts',JSON.stringify(counts));
 render();
 qty.value='';
}

function render(){
 list.innerHTML='<tr><th>کد</th><th>نام</th><th>مقدار</th></tr>';
 counts.forEach(x=>{
  list.innerHTML+=`<tr><td>${x.code}</td><td>${x.name}</td><td>${x.qty}</td></tr>`;
 });
}

function exportExcel(){
 const rows=counts.map(x=>({
 'نوع قلم':'InventoryDeliveryItem',
 'نوع قلم خروج انبار':4,
 'نوع خروج انبار':1531,
 'تاریخ خروج انبار':invDate.value,
 'کد انبار':sourceWarehouse.value,
 'کد کالا':x.code,
 'مقدار اصلی':x.qty,
 'کد حساب معین':111509,
 'کد انبار مقصد':destWarehouse.value
 }));
 const ws=XLSX.utils.json_to_sheet(rows);
 const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,ws,'Sepidar');
 XLSX.writeFile(wb,'InventoryDelivery.xlsx');
}
render();
