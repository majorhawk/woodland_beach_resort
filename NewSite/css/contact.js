
function checkFields(form) {
	
	form.FNme.style.border = "1px solid #BFB18E";
	form.LNme.style.border = "1px solid #BFB18E";
	form.Email.style.border = "1px solid #BFB18E";
	
	
// First Name Validation Statement
	if (form.FNme.value == ""){
			alert('Please Enter Your First Name');
			form.FNme.focus();
			form.FNme.style.border = "2px solid #CC0000";
			return false;
		}
	
// Last Name Validation Statement
	if (form.LNme.value == ""){
			alert('Please Enter Your Last Name');
			form.LNme.focus();
			form.LNme.style.border = "2px solid #CC0000";
			return false;
		}
	
	
// Email Address Validation Statement
	if (form.Email.value == ''){
			
			alert("Please Enter Your Email Address");
			form.Email.focus();
			form.Email.style.border = "2px solid #CC0000";
			return false;
		}
			
	if (form.Email.value.indexOf('@') == -1){
			
			alert("Please Enter a Valid Email Address That Contains @");
			form.Email.focus();
			form.Email.style.border = "2px solid #CC0000";
			return false;
		}
			
	if (form.Email.value.indexOf('.') == -1){
			
			alert("Please Enter a Valid Email Address That Contains a Valid Extension.\nexample: .com, .net, .org, .biz, .edu, .gov");
			form.Email.focus();
			form.Email.style.border = "2px solid #CC0000";
			return false;
		}
		
}