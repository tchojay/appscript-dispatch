function onFormSubmit(e){

  //Dispatch No generation
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var lastRow = spreadsheet.getLastRow();
  var dispatch_arr = spreadsheet.getRange("A2:A").getValues();
  var arr_to_num = [].concat.apply([],dispatch_arr);
  var max_disp = Math.max.apply(Math,arr_to_num);
  var disNo = max_disp +1;


  // Form Fields
  var formValues = e.namedValues;
  
  //HTML Body

    var html = "<br>Dear "+ formValues['Email Address'] + ",<h3> Letter Dispatch No requested is:: "+disNo +"</h3> <br> Following are the details of the requested letter dispatch:<br><br>";
 
  for(key in formValues){
    Logger.log (key + ' '+ formValues[key]);
    html+= '<div><b>'+key+' '+'</b>:'+formValues[key] +'</div>';
    
  }
  html+= "<br><br><div style='text-align:center; font-style:italic;'> Disclaimer: This letter dispatch no. is system generated and it is intended to be used only for letters signed officially. </div> <div style='text-align:center'><hr><h4> Request for another letter dispatch no. <a href='https://forms.gle/fFQnx2gvE7S71nLD8' target='_blank'>Click Here</a> </h4><h6> Powered by G Suite Team, DITT | support@gov.bt </h6><hr><br></div> ";
  
  
  //Send Email
  var email = formValues["Email Address"].toString();  
  MailApp.sendEmail(email, 'Letter Dispatch No','',{
    htmlBody: html});
  spreadsheet.getRange(lastRow, 1).setValue(disNo);
  
}

