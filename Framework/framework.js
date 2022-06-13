function setCookie(cname, cvalue, exdays) {
    var d = new Date();
    d.setTime(d.getTime() + (exdays*24*60*60*1000));
    var expires = "expires="+d.toUTCString();
    document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/";
}
function getCookie(cname) {
    var name = cname + "=";
    var ca = document.cookie.split(';');
    for(var i = 0; i < ca.length; i++) {
        var c = ca[i];
        while (c.charAt(0) == ' ') {
            c = c.substring(1);
        }
        if (c.indexOf(name) == 0) {
            return c.substring(name.length, c.length);
        }
    }
    return "";
}
var formatMoney = function(n, c, d, t){
	var c = isNaN(c = Math.abs(c)) ? 2 : c, 
		d = d == undefined ? "," : d, 
		t = t == undefined ? "." : t, 
		s = n < 0 ? "-" : "", 
		i = parseInt(n = Math.abs(+n || 0).toFixed(c)) + "", 
		j = (j = i.length) > 3 ? j % 3 : 0;
	   return s + (j ? i.substr(0, j) + t : "") + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + t) + (c ? d + Math.abs(n - i).toFixed(c).slice(2) : "");
}; 
function getLinkBySystem(link){
	var system = getParameterByName("system");
	
	if(link!=undefined && link!="#" && link.search("pg")==-1 && link.search("system")==-1){
		if(system!="" && system!=null){
			link+="&pg=redirect&system="+system;
		}	
	}
	return link;
}
function getParameterByName(name) {
    url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}

function copyToClipboard(value) {
	var success   = true,
	  range     = document.createRange(),
	  selection;

	// For IE.
	if (window.clipboardData) {
		window.clipboardData.setData("Text", value);        
	} else {
		// Create a temporary element off screen.
		var tmpElem = $('<div>');
		tmpElem.css({
		  position: "absolute",
		  left:     "-1000px",
		  top:      "-1000px",
		});
		// Add the input value to the temp element.
		tmpElem.text(value);
		$("body").append(tmpElem);
		// Select temp element.
		range.selectNodeContents(tmpElem.get(0));
		selection = window.getSelection ();
		selection.removeAllRanges ();
		selection.addRange (range);
		// Lets copy.
		try { 
			success = document.execCommand ("copy", false, null);
		}
		catch (e) {
			alert("Copiar para a área de tansferência não suportada neste navegador.")
		}
		if (success) {
			alert ("Caminho copiado com sucesso.");
			// remove temp element.
			tmpElem.remove();
		}
	}
}


function alert(string, type, nome, autoclose){
	console.log(string);
	var nomeF;
	if(type==undefined){
		if(typeof string == "string"){
			if(string.search("sucess")!=-1 || string.search("succes")!=-1){
				type = "success";
			}else if(string.search("erro")!=-1){
				type = "danger";
			}else{
				type="info";
			}
		}
	}
	if(nome!=undefined && nome!=""){
		nomeF=nome;
	}else{ 
		var numero=Math.floor((Math.random() * 1000) + 1);
		nomeF="alert"+numero;
	}
	if(type=="danger-auto"){
		close=string;
		type="danger";	
		if(autoclose==undefined){
			autoclose=8000;
		}
	}else if(type=="error"){
		close=string;
		type="danger";
		if(autoclose==undefined){
			autoclose=8000;
		}
	
	}else if(type=="loading"){
		close=string+'<span class="loadingGif">&nbsp;</span>';
	}else{
		if(typeof string == "string"){
			if(string.search("sucess")!=-1 || string.search("succes")!=-1){
				close = string;
			}else if(string.search("erro")!=-1){
				close = string;
			}else{
				close = string;
				//close='<button type="button" class="close"><span aria-hidden="true">×</span><span class="sr-only">Close</span></button>'+string;
			}
		}else{
			close = string;
			//close='<button type="button" class="close"><span aria-hidden="true">×</span><span class="sr-only">Close</span></button>'+string;
		}
		if(autoclose==undefined){
			autoclose=8000;
		}
	}
	var el = $(`<div class="pro-banner" id="`+nomeF+`">
				<div class="card pro-banner-bg border-0 rounded-0">
					<div class="card-body py-3 px-4 d-flex align-items-center justify-content-between flex-wrap">
						<p class="mb-0 text-white font-weight-medium mb-2 mb-lg-0 mb-xl-0" >`+close+`</p>
						<div class="d-flex">
							<button id="bannerClose" class="close">
								<i class="mdi mdi-close text-white"></i>
							</button>
						</div>
					</div>
				</div>
				</div>`);
	//var el = $('<div class="alert-top alert alert-'+type+' page-alert" id="'+nomeF+'">'+close+'</div>');
	$('.page-alerts').append(el);
	
    var timeOut;
	
	var alerta=$("#"+nomeF);

	alerta.find('.close').click(function(e) {
        e.preventDefault();
        //$(this).closest('.page-alert').slideUp();
		$(this).closest('.pro-banner').slideUp();
    });
	
	alerta.hide();
    alerta.slideDown();
   
   
   if(autoclose!=undefined){
		var delay = autoclose;
			
		delay = parseInt(delay);
		clearTimeout(timeOut);
		timeOut = window.setTimeout(function() {
				alerta.slideUp();
			}, delay);
	}
}

TipoCampo = {
    DATE : 0,
    HOUR : 1,
    PLATE : 2,
    HOURANDMINUTES : 3,
    CNPJ : 4,
}
var validaData = function (str){
	//Split data.
	var data_array = str.split('/');
	//alert(document.getElementById(id).value);
	var dia = data_array[0];
	//Javascript considera meses de 0 - 11
	var mes = data_array[1] - 1;
	var ano = data_array[2];
	if(ano<1900)
		return false;
	//Criando um objeto data
	var dataSistema = new Date(ano,mes,dia);
	/*dataSistema.setDate(dia)
	dataSistema.setMonth(mes);
	dataSistema.setYear(ano);*/
	//Compara data digitada com data do objeto Date do Javascript.
	if (ano != dataSistema.getFullYear() ||  mes != dataSistema.getMonth() ||dia != dataSistema.getDate() ){
		return false;
	}
	return true;
}

var verifyDataField = function(field){
	if(checkForm(field.val(), TipoCampo.DATE) && validaData(field.val())){
		field.parent().removeClass("has-warning").removeClass("has-error").removeClass("has-success");
		return true;
	}else{
		field.parent().removeClass("has-warning").removeClass("has-success").addClass("has-error");
		return false;
	}
}
var verifyPlateField = function(field){
	if(checkForm(field.val(), TipoCampo.PLATE)){
		field.parent().removeClass("has-warning").removeClass("has-error").removeClass("has-success");
		return true;
	}else{
		field.parent().removeClass("has-warning").removeClass("has-success").addClass("has-error");
		return false;
	}
}
var verifyCA = function(e){	
	if (e == "") {
		alert("Informe o CA.")
		return false;
	} else if (isNaN(e)) {
		alert("CA inválido: "+e)
		return false;
	} else {	
		return true;
	}
}

var verifyCPFField = function(field){
	if(field.val().length>13){
		if(checkForm(field.val(), TipoCampo.CPF) && validarCPF(field.val())){
			field.parent().removeClass("has-warning").removeClass("has-error").removeClass("has-success");
			return true;
		}else{
			field.parent().removeClass("has-warning").removeClass("has-success").addClass("has-error");
			return false;
		}
	}else{
		return false;
	}
}
var verifyCNPJField = function(field){
	if(checkForm(field.val(), TipoCampo.CNPJ) && validarCNPJ(field.val())){
		field.parent().removeClass("has-warning").removeClass("has-error").removeClass("has-success");
		return true;
	}else{
		field.parent().removeClass("has-warning").removeClass("has-success").addClass("has-error");
		return false;
	}
}	
var verifyHoraField = function(field){
	//if(checkForm(field.val(), TipoCampo.HOUR)){
		if($.isNumeric(field.val().split(" ")[0])){
			field.parent().removeClass("has-warning").removeClass("has-error").removeClass("has-success");
			return true;
		}else{
			field.parent().removeClass("has-warning").removeClass("has-success").addClass("has-error");
			return false;
		}
	//}else{
	//	field.parent().removeClass("has-warning").removeClass("has-success").addClass("has-error");
	//	return false;
	//}
}
var verifyHoraEMinutosField = function(field){
	if(checkForm(field.val(), TipoCampo.HOURANDMINUTES)){
		let temp = field.val();
		let temp2 = temp.toString();
		if( field.val().indexOf(" ") === -1){
			field.parent().removeClass("has-warning").removeClass("has-error").removeClass("has-success");
			return true;
		}
		else 
		{
			if($.isNumeric(temp2.split(" ")[0])){
				field.parent().removeClass("has-warning").removeClass("has-error").removeClass("has-success");
				return true;
			}else{
				field.parent().removeClass("has-warning").removeClass("has-success").addClass("has-error");
				return false;
			}
		}
	}else{
		field.parent().removeClass("has-warning").removeClass("has-success").addClass("has-error");
		return false;
	}
}

//Converte data de DD/MM/AAAA para AAAA-MM-DD e vice versa
var dateFormat = function(data){
	var V_data ="";
	var A;
	if(data!=undefined && data.length>0){
		if (data.search("/")>=0){
			A = data.split ("/");
			V_data = A[2]+"-"+A[1]+"-"+A[0];
		}else if (data.search("-")){
			A = data.split ("-");
			V_data = A[2]+"/"+A[1]+"/"+A[0];
		}else{
			V_data = "";
		}
	}
	return V_data;
}
//Regex para substituir primeira letra de cada palavra por caps
function toTitleCase(str)
{
    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}
var mascara = function (o,f){
	v_obj=o
	v_fun=f
	setTimeout("execmascara()",1)
}

var execmascara = function(){
	v_obj.val(v_fun(v_obj.val()));
}
//Formata data 12/12/1234
var mdata = function (v){
	v=v.replace(/\D/g,"");                    
	v=v.replace(/(\d{2})(\d)/,"$1/$2");       
	v=v.replace(/(\d{2})(\d)/,"$1/$2");       

	v=v.replace(/(\d{2})(\d{2})$/,"$1$2");
	return v;
}
//Formata placa de carro DDD-0000
var mplate = function (v){  
	v=v.toUpperCase();
	v=v.replace(/\W+/,"");  
	v=v.replace(/([A-Z]{3})(\d{4})/,"$1-$2");
	return v;
}
//Somente numeros
var mnumbers = function (v){  
	v=v.replace(/\D/,"");  
	return v;
}
//Transforma numeros em valores
var mvalue = function (v){  
	if (v.indexOf(".") >= 0 && v.indexOf(",") >= 0){
		v=v.replace(".","");
		v=v.replace(",",".");
	}

	v=v.replace(",","."); 
	v = formatMoney(v, 2);
	return v;
}
//retirado do portal
//Formata o 00000-000
var mcep = function (v){  
	v=v.toUpperCase();
    v=v.replace(/\D/g,"");
	v=v.replace(/(\d{5})(\d{3})/,"$1-$2");
	return v;
}
//Strings Sao Digitadas Assim
var mletras = function (v){  
	v=toTitleCase(v); 
	return v;
}
//STRINGS SAO DIGITADAS ASSIM
var mallcaps = function (v){  
	v=v.toUpperCase(); 
	return v;
}

//Strings Sao Digitadas Assim
var lettersnumbers = function (v){  
	v=v.toUpperCase(); 
	return v;
}

//Formata strings (00) 00000-0000
var mtel = function(v){
    v=v.replace(/\D/g,"");             //Remove tudo o que não é dígito
    v=v.replace(/^(\d{2})(\d)/g,"($1) $2"); //Coloca parênteses em volta dos dois primeiros dígitos
    v=v.replace(/(\d)(\d{4})$/,"$1-$2");    //Coloca hífen entre o quarto e o quinto dígitos
    return v;
}
//Formata 23 H
var mhora = function (v){ 
	v=v.replace(/\D/g,"");       
	v=v.replace(/(\d{2})(\d)/,"$1"); 
	if(parseInt(v)>23){
		v="23";
	}else if(parseInt(v)<0){
		v="00";
	}
	v=v+" H";
	return v;
}	
//Formata 23:00
var mhoraeminutos = function (v){ 
	v=v.replace(/\D/g,"");       
	v=v.replace(/^(\d{2})(\d)/g,"$1:$2"); 
	
	if(v.length==5){
		var test = v.split(":");
		if(test[0]>23){
			test[0]="23";
		}else if(test[0]<0){
			v="00";
		}
		if(test[1]>60){
			test[1]="59";
		}else if(test[1]<0){
			v="00";
		}
		v = test[0]+":"+test[1];
		}
	return v;
}	
//Formata CPF 000.000.000-00
var mcpf = function (v){ 
	v=v.replace(/\D/g,"");       
	v=v.replace(/(\d{3})(\d)/,"$1.$2"); 
	v=v.replace(/(\d{3})(\d)/,"$1.$2"); 
	v=v.replace(/(\d{3})(\d)/,"$1-$2"); 
	return v;
}	
//Formata CNPJ 00.000.000/0000-00
var mcnpj = function (v){	
	v=v.replace(/\D/g,"");       
	v=v.replace(/(\d{2})(\d)/,"$1.$2"); 
	v=v.replace(/(\d{3})(\d)/,"$1.$2"); 
	v=v.replace(/(\d{3})(\d)/,"$1/$2"); 
	v=v.replace(/(\d{4})(\d)/,"$1-$2"); 
	return v;
}	
//Formata RG 00.000.000
var mrg = function (v){ 
	//Carlos [08/03/2016 16:50]: na verdade nao existe essa regra, pois rg realmente é sem padrao
	//
	v=v.toUpperCase();
	v=v.replace(/\D/g,"");   
	/*v=v.replace(/(\d{3})(\d)/,"$1.$2"); 
	v=v.replace(/(\d{3})(\d)/,"$1.$2"); 
	v=v.replace(/(\d{3})(\d)/,"$1.$2"); 
	v=v.replace(/(\d{3})(\d)/,"$1.$2"); */
	return v;
}

var validarCNPJ = function(cnpj) {
	cnpj = cnpj.replace(/\./g, "").replace(/\-/g, "").replace(/\//g, "");
	cnpj = cnpj.replace(/[^\d]+/g,'');

	if(cnpj == '') 
		return false;

	if (cnpj.length != 14)
		return false;

	// Elimina CNPJs invalidos conhecidos
	if (cnpj == "00000000000000" || 
	cnpj == "11111111111111" || 
	cnpj == "22222222222222" || 
	cnpj == "33333333333333" || 
	cnpj == "44444444444444" || 
	cnpj == "55555555555555" || 
	cnpj == "66666666666666" || 
	cnpj == "77777777777777" || 
	cnpj == "88888888888888" || 
	cnpj == "99999999999999")
		return false;

	// Valida DVs
	tamanho = cnpj.length - 2
	numeros = cnpj.substring(0,tamanho);
	digitos = cnpj.substring(tamanho);
	soma = 0;
	pos = tamanho - 7;
	for (i = tamanho; i >= 1; i--) {
		soma += numeros.charAt(tamanho - i) * pos--;
		if (pos < 2)
		pos = 9;
	}
	resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
	if (resultado != digitos.charAt(0))
		return false;

	tamanho = tamanho + 1;
	numeros = cnpj.substring(0,tamanho);
	soma = 0;
	pos = tamanho - 7;
	for (i = tamanho; i >= 1; i--) {
		soma += numeros.charAt(tamanho - i) * pos--;
		if (pos < 2)
			pos = 9;
	}
	resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
	if (resultado != digitos.charAt(1))
		return false;
	       
	return true;
}

var validarCPF  = function(cpf) {
	cpf = cpf.replace(/\./g, '').replace(/\-/g, '');
	if (cpf.length != 11 || cpf == "00000000000" || cpf == "11111111111" || cpf == "22222222222" || cpf == "33333333333" || 
		cpf == "44444444444" || cpf == "55555555555" || cpf == "66666666666" || cpf == "77777777777" || cpf == "88888888888" || 
		cpf == "99999999999")
		return false;
	
    var Soma;
    var Resto;
    Soma = 0;
	if (cpf == "00000000000") return false;
    
	for (i=1; i<=9; i++) Soma = Soma + parseInt(cpf.substring(i-1, i)) * (11 - i);
	Resto = (Soma * 10) % 11;
	
	
    if ((Resto == 10) || (Resto == 11))  Resto = 0;
    if (Resto != parseInt(cpf.substring(9, 10)) ) return false;
	
	Soma = 0;
    for (i = 1; i <= 10; i++) Soma = Soma + parseInt(cpf.substring(i-1, i)) * (12 - i);
    Resto = (Soma * 10) % 11;

	
	
    if ((Resto == 10) || (Resto == 11))  Resto = 0;
    if (Resto != parseInt(cpf.substring(10, 11) ) ) return false;
	
    return true;
}

//End Portal
var checkForm = function(text, tipo){
	if(tipo==TipoCampo.DATE)
		re = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
	else if(tipo==TipoCampo.HOURANDMINUTES)
		re = /^\d{1,2}:\d{2}([ap]m)?$/;
	else if(tipo==TipoCampo.PLATE)
		re = /^([A-Z]{3})-(\d{4})([ap]m)?$/;
	else if(tipo==TipoCampo.CPF)
		re = /^(\d{3}).(\d{3}).(\d{3})-(\d{2})([ap]m)?$/;
	else if(tipo==TipoCampo.CNPJ)
		//17.320.401/0001-21
		re = /^(\d{2}).(\d{3}).(\d{3})\/(\d{4})-(\d{2})([ap]m)?$/;
	else 
		return false;	

	if(text == '' || ((text+"").match(re)==null || (text+"").match(re)==undefined || (text+"").match(re)=="")) {
		return false;
	}

	return true;
 }
 //Retirado do Portal Colaborador
 //Checa se o CPF é válido
var TestaCPF = function(strCPF) {
	strCPF=strCPF.replace(/\./g, "");
	strCPF=strCPF.replace("-", "");
	//alert(strCPF);	
	var Soma;
	var Resto;
	Soma = 0;
	if (strCPF == "00000000000")
		return false;
	for (i = 1; i <= 9; i++)
		Soma = Soma + parseInt(strCPF.substring(i - 1, i)) * (11 - i);
	Resto = (Soma * 10) % 11;
	if ((Resto == 10) || (Resto == 11))
		Resto = 0;
	if (Resto != parseInt(strCPF.substring(9, 10)))
		return false;
	Soma = 0;
	for (i = 1; i <= 10; i++)
		Soma = Soma + parseInt(strCPF.substring(i - 1, i)) * (12 - i);
	Resto = (Soma * 10) % 11;
	if ((Resto == 10) || (Resto == 11))
		Resto = 0;
	if (Resto != parseInt(strCPF.substring(10, 11)))
		return false;
	return true;
}
 //
 var sendDefaultForm = function(element, extraCheck){
	var podeIr=true;
	var alerta=true;
	
	element.find("input, textarea").each(function(){
		var attr = $(this).attr('temp');
		var thisPode=true;

		if (typeof attr !== typeof undefined && attr !== false) {
			if($(this).attr("temp")=="true"){
				$(this).parent().remove();
			}
		}
		
		
		if($(this).attr("required")=="required"){
			if($(this).val()!=""){
				switch($(this).attr("input-type")){
					case "horaeminutos":
						if(!verifyHoraEMinutosField($(this))){							
							thisPode=false;
							break;
						}
					break;
					case "hora":
						if(!verifyHoraField($(this))){							
							thisPode=false;
							break;
						}
					break;
					case "date":
						if(!verifyDataField($(this))){							
							thisPode=false;
							break;
						}
					break;
					case "plate":
						if(!verifyPlateField($(this))){
							thisPode=false;
							break;
						}
					break;	
					case "ca":
						if(!verifyCA($(this).val())){
							thisPode=false;
							break;
						}
					break;
					case "cpf":
						if(!verifyCPFField($(this))){
							thisPode=false;
							break;
						}
					break;
					case "cnpj":
						if(!verifyCNPJField($(this))){
							thisPode=false;
							break;
						}
					break;
					default:					
						//$(this).parent().removeClass("has-warning").removeClass("has-error").addClass("has-success");
					break;
				}
			}else{
				thisPode=false;
			}
		}else{
			$(this).parent().removeClass("has-warning").removeClass("has-error");
		}
		
		
		//Depois de todos os checks, liberar
		if(thisPode){		
			$(this).parent(".form-group").first().removeClass("has-warning").removeClass("has-error").addClass("has-success");
		}else{
			podeIr=false;
			$(this).parent(".form-group").first().removeClass("has-warning").removeClass("has-success").addClass("has-error");
		}
	});
	if(podeIr){	
		if(extraCheck!=undefined){
			if(!extraCheck()){
				return false;
			}
		}

		element.parent("form").submit();
		return true;
	}else{
		//if(alerta){
			alert("Por favor, preencha todos os campos em vermelho (obrigatório).", "danger");
		//}
		return false;
	}
	return false;
}

var stripHTML = function strip(html){
   var tmp = document.createElement("DIV");
   tmp.innerHTML = html;
   return tmp.textContent||tmp.innerText;
}

var intializeSelects = function(){	
	$("select").each(function(){
		var count = $(this).find("option").length;
		var placeholder = $(this).attr("inputholder");
		if(placeholder != undefined){
			$(this).attr("data-placeholder",placeholder);
		}
		console.log(placeholder);
		if(count>1){
			$(this).chosen({
				placeholder_text_multiple: placeholder,
				placeholder_text_single: "Escolha uma opção",
				search_contains: true,
				disable_search_threshold: 10,
				no_results_text: "Nenhum resultado encontrado.",
				width: "100%"
			});
		}
	});
}
var DataTables, DataTables2 = "";

var initializeDataTables = function(){
	var languageData = {
			"emptyTable":     "Nenhum dado disponível",
            'sLengthMenu': 'Mostrar _MENU_ registros por página',
            "lengthMenu": "Mostrando _MENU_ resultados por pagina",
            "zeroRecords": "Nenhum resultado encontrado.",
            "info": "Mostrando página _PAGE_ de _PAGES_ de _TOTAL_ registros",
            "infoEmpty": "Nenhum resultado encontrado",
            "infoFiltered": "(filtrando de _MAX_ resultados)",
			"processing":     "Carregando...",
			"search": "Buscar:" 
	};
			
	var buttonsData = ['excelHtml5'];
	var processing = false, serverSide = false, ajaxPath="";
	var processing2 = false, serverSide2 = false, ajaxPath2="";
	
	if($('#dataTables').attr("ajax")!=undefined){
		processing = true;
		serverSide = true;
		ajaxPath = {
			"url": $('#dataTables').attr("ajaxpath"),
			"type": "POST"
		};
	}
	if($('#dataTables2').attr("ajax")!=undefined){
		processing2 = true;
		serverSide2 = true;
		ajaxPath2 = {
			"url": $('#dataTables2').attr("ajaxpath2"),
			"type": "POST"
		};
	}
	
	if($('#dataTables').length){
		DataTables = $('#dataTables').DataTable({
			dom: '<"top"Blf>rt<"bottom"ip><"clear">',
			buttons: buttonsData,
			"pageLength":10,
			responsive: true,
			"language": languageData,
			//ajax
			"processing": processing,
			"serverSide": serverSide,
			"ajax": ajaxPath
		});
	}
	if($('#dataTables2').length){
		DataTables2 = $('#dataTables2').DataTable({
			dom: 'Bf<"bottom"l>rtip',
			buttons: buttonsData,
			"pageLength":10,
			responsive: true,
			"language": languageData,
			//ajax
			"processing": processing2,
			"serverSide": serverSide2,
			"ajax": ajaxPath2
		});
	}
}

var Notifications = {};
Notifications.Request = function(){
	if(!window.Notification) {
		console.log('Este browser não suporta Web Notifications!');
		return false;
	}
	if(Notification.permission === 'default'){
		Notification.requestPermission(function(){
			console.log('No permission found. Requesting again.');
			return false;
		});
	}else if(Notification.permission === 'granted'){
		//console.log("Permission already granted");
		return true;
	} else if (Notification.permission === 'denied') {
		console.log('Usuário não deu permissão');
		return false;
	}
}
//Notifications.Send = function(title, text, onClick=null, onClose=null){
Notifications.Send = function(title, text, onClick, onClose){
	var folder = window.location.href.split("/")[3];
	var derp = new Notification(title, { 
			body: text,
			icon: './'+folder+'/graphics/images/roleptLogo.png',
			tag: 'Mais de uma notificação para o Sistema Administrativo',
		});
		
	derp.onshow = function() {
		console.log('onshow: evento quando a notificação é exibida')
	},
	derp.onclick = function() {
		if(onClick != undefined && onClick!=null){
			onClick();
		}
		derp.close();
	},
	derp.onclose = function() {
		console.log('onclose: evento quando a notificação é fechada')
	},
	derp.onerror = function() {
		console.log('onerror: evento quando a notificação não pode ser exibida. É disparado quando a permissão é defualt ou denied')
	}
}

	
$(document).on("click", "a", function(e){
	
	var system = getParameterByName("system");
	var isGlobalLink = $(this).attr("isGlobalLink");
	var inputType = $(this).attr("input-type");
	var tagIt = $(this).hasClass("tagit-label");
	
	//var href = $(this).attr("href").search("#");
	
	var href = $(this).attr("href");	
	if(!(href.includes("#"))){
		if(system!="" && system!=null && (isGlobalLink==null || isGlobalLink=="") && (inputType==null || inputType=="") && tagIt==false){	
			href+="&pg=redirect&system="+system;	
			$(this).attr("href", href);	
			//return true;	
		}
	}else{	
		e.preventDefault();
		return false;										
	}			
	
	return true;
	
	/*
	if(href<0){
		alert("1");
		if(system!="" && system!=null && (isGlobalLink==null || isGlobalLink=="") && (inputType==null || inputType=="") && tagIt==false){
			href+="&pg=redirect&system="+system;
			$(this).attr("href", href);
			//return true;
		}	
	}
	e.preventDefault();
	return false;*/
});

$(document).on("click", "a[input-type='edit']", function(event){
	var form = $(this).parents("form");
	var action = form.attr("action");
	form.attr("action", getLinkBySystem(action));
	
	

	$(this).parent("form").submit();
	event.preventDefault();
});

$(document).on("click", "form button[input-type='submit']", function(event){
	var form = $(this).parents("form");
	var action = form.attr("action");
	form.attr("action", getLinkBySystem(action));

	//alert(form.attr("action"));
	return sendDefaultForm(form);
	event.preventDefault();
})
	
// $( document ).ready(function(){
// 	Notifications.Request();
	
	
// 	//Inicializa possiveis DataTables
// 	intializeSelects();
// 	initializeDataTables();
	
// 	$.datepicker.regional["pt-BR"];
// 	$('body').on('focus', "input[input-type='date']", function(){
// 		$(this).datepicker({
// 		  changeMonth: true,
// 		  changeYear: true,
// 		  showOtherMonths: true,
// 		  selectOtherMonths: true
// 		});
// 	});
// 	$(document).on("keypress", "input[input-type='date']", function(){
// 		$(this).attr('maxlength','10');
// 		mascara($(this), mdata);
// 		verifyDataField($(this));
// 	});
// 	$(document).on("keyup", "input[input-type='date']", function(){
// 		$(this).attr('maxlength','10');
// 		mascara($(this), mdata);
// 		verifyDataField($(this));
// 	});
	
	
	
	
// 	$(document).on("keypress", "input[input-type='plate']", function(){
// 		$(this).attr('maxlength','8');
// 		mascara($(this), mplate);
// 		verifyPlateField($(this));
// 	});
// 	$(document).on("keypress", "input[input-type='number']", function(){
// 		mascara($(this), mnumbers);
// 	});
// 	//Tirado do Portal	
// 	//Formata campos de input para cep
// 	$(document).on("keypress", "input[input-type='cep']", function(){
// 		$(this).attr('maxlength','9');
// 		updateCEP($(this));
// 	});
// 	$(document).on("change", "input[input-type='cep']", function(){
// 		$(this).attr('maxlength','9');
// 		updateCEP($(this));
// 	});
// 	var updateCEP = function(who){
// 		mascara(who, mcep);
// 	}
// 	//Formata campos de input para Deste Jeito
// 	$(document).on("blur", "input[input-type='letters']", function(){
// 		updateLetters($(this));
// 	});
// 	$(document).on("change", "input[input-type='letters']", function(){
// 		updateLetters($(this));
// 	});
// 	$(document).on("keypress", "input[input-type='letters']", function(){
// 		updateLetters($(this));
// 	});
// 	var updateLetters = function(who){
// 		mascara(who, mletras);
// 	}
	
// 	//Formata os campos de input para 100.000.000,00
// 	$(document).on("blur", "input[input-type='value']", function(){
// 		updateValue($(this));
// 	});
// 	var updateValue = function(input){
// 		mascara(input, mvalue);
// 	}
// 	//FORMATA OS CAMPOS DE INPUT DESDE JEITO
// 	$(document).on("keypress", "input[input-type='allcaps']", function(){
// 		updateAllCaps($(this));
// 	});
// 	var updateAllCaps = function(who){
// 		mascara(who, mallcaps);
// 	}
	
	
// 	//Formata campos de input para XX-12.345.678
// 	$(document).on("keypress", "input[input-type='rg']", function(){
// 		updateRG($(this));
// 	});
// 	$(document).on("blur", "input[input-type='rg']", function(){
// 		updateRG($(this));
// 	});
// 	$(document).on("change", "input[input-type='rg']", function(){
// 		updateRG($(this));
// 	});
// 	var updateRG = function(who){
// 		mascara(who, mrg);
// 	}
	
	
// 	//Formata campos de input para (xx) 00000-0000
// 	$(document).on("keypress", "input[input-type='telefone']", function(){
// 		$(this).attr('maxlength','15');
// 		updateTel($(this));
// 	});
// 	$(document).on("change", "input[input-type='telefone']", function(){
// 		$(this).attr('maxlength','15');
// 		updateTel($(this));
// 	});
// 	var updateTel = function(who){
// 		mascara(who, mtel);
// 	}
	
	
// 	//Formata campos de input para 00.0000.000/0000-000
// 	$(document).on("keypress", "input[input-type='cnpj']", function(){
// 		$(this).attr('maxlength','18');
// 		mascara($(this), mcnpj);
// 		verifyCNPJField($(this));
// 	});		
// 	$(document).on("change", "input[input-type='cnpj']", function(){
// 		$(this).attr('maxlength','18');
// 		mascara($(this), mcnpj);
// 		verifyCNPJField($(this));
// 	});		
// 	//Formata campos de input para 000.000.000-00
// 	$(document).on("keypress", "input[input-type='cpf']", function(){
// 		$(this).attr('maxlength','14');
// 		mascara($(this), mcpf);
// 		verifyCPFField($(this));
// 	});		
// 	$(document).on("keyup", "input[input-type='cpf']", function(){
// 		$(this).attr('maxlength','14');
// 		mascara($(this), mcpf);
// 		verifyCPFField($(this));
// 	});		
// 	//Formata campos de input para 000.000.000-00
// 	$(document).on("change", "input[input-type='cpf']", function(){
// 		$(this).attr('maxlength','14');
// 		mascara($(this), mcpf);
// 		verifyCPFField($(this));
// 	});		
	
// 	//Formata campos de input para XX H
// 	$(document).on("keypress", "input[input-type='hora']", function(){
// 		$(this).attr('maxlength','4');
// 		mascara($(this), mhora);
// 		verifyHoraField($(this));
// 	});		
// 	$(document).on("change", "input[input-type='hora']", function(){
// 		$(this).attr('maxlength','4');
// 		mascara($(this), mhora);
// 		verifyHoraField($(this));
// 	});		

// 	//end
	
// 	//Formata campos de input para HH:MM
// 	$(document).on("keypress", "input[input-type='horaeminutos']", function(){
// 		$(this).attr('maxlength','5');
// 		mascara($(this), mhoraeminutos);
// 		verifyHoraEMinutosField($(this));
// 	});	
	
// 	//end
	

// 	var $_GET = {};

// 	document.location.search.replace(/\??(?:([^=]+)=([^&]*)&?)/g, function () {
// 		function decode(s) {
// 			return decodeURIComponent(s.split("+").join(" "));
// 		}

// 		$_GET[decode(arguments[1])] = decode(arguments[2]);
// 	});
// 	$(document).on("click", "a[input-type='erase']", function(){
// 		var id = $(this).attr("input-id");
// 		var me = $(this);
// 		var link = $(this).attr("href");
// 		var st = link.split("status=");

// 		status = st[1];
		
// 		swal({
// 		  title: "Confirmar exclusão",
// 		  text: "Este item será excluido",
// 		  type: "error",
// 		  showCancelButton: true,
// 		  confirmButtonClass: "btn-danger",
// 		  confirmButtonText: "Apagar",
// 		  closeOnConfirm: false
// 		}).then(function(inputValue){
// 			$.ajax({	
// 				method: "POST",
// 				url: link,//"?pagina="+getParameterByName("pagina")+"&noHeader=true&status="+status,
// 				data: { id: id, status: status}
// 			})
// 			.done(function(msg) {
// 				msg = stripHTML(msg);
// 				if(msg=="1"){
// 					swal("Excluido!", "Este ítem foi excluido.", "success");
// 					me.parent().parent().addClass("selected");
// 					var dt;
// 					if(me.parents("#dataTables").length){
// 						dt = $('#dataTables').DataTable();
// 					}else{
// 						dt = $('#dataTables2').DataTable();
// 					}
// 					dt.row('.selected').remove().draw( false );
// 				}else{
// 					swal("Erro!", "Ocorreu um erro, por favor contate o suporte do sistema.", "warning");
// 				}
// 			});
			
// 		}, function(){});
// 		event.preventDefault();
// 	});
// 	$(document).on("change", "select", function(){
// 		var name = $(this).attr("id").replace("_select", "");
// 		$(this).parent().find("input[name='"+name+"']").val($(this).val());
// 	});
// 	$(document).on("change", "input[type='checkbox']", function(){ 
// 		if($(this).prop("checked")){
// 			$(this).next().val(1);
// 		}else{
// 			$(this).next().val(0);
// 		}
// 	});
// 	$(document).on("click", "a[class='downcontent']", function(){
// 		var offset = $(this).offset();
// 		var hasdialog = $(this).attr("has-dialog");
// 		var unique = $(this).attr("unique");
// 		var lastdialog = $(this).attr("last-dialog");
// 		var titledialog = $(this).attr("title-dialog");
// 		var textdialog = $(this).attr("text-dialog");
// 		var id = "1";
// 		id = id.replace("downcontent", "");
// 		var background = $(this).css("background-color");
// 		var subtraioffset = 0;
// 		var dialogposition = 0;
// 		var unique1 = $(this).attr("unique");
// 		var unique2 = $(".speech-bubble");
// 		var index = 0
		
// 		//tratamento p/ quando tem mais de 1 time line na msm pagina
// 		for(var i = 0; i < $(".titledialog").length; i++){
// 			if(unique1 == unique2[i].getAttribute("unique")){			
// 				index = i
// 			}
// 		}

// 		if(lastdialog == 1){
// 			subtraioffset = 140;
// 			dialogposition = "61%";
// 		}else{
// 			subtraioffset = 5
// 			dialogposition = "17%";
// 		}
// 		if(hasdialog == 1){
// 			if($(".dropup-content"+id).get(index).style.getPropertyValue("display") != "none"){
// 				$(".dropup-content"+id).get(index).style.setProperty("display", "none");						
// 			}else{
// 				$(".speech-bubble").get(index).style.setProperty("--color", background);	
// 				$(".titledialog").get(index).innerHTML = titledialog;				
// 				$(".textdialog").get(index).innerHTML = textdialog;				
// 				$(".dropup-content"+id).get(index).style.setProperty("display", "block");
// 				//$(".dropup").offset({left: offset.left - subtraioffset});	
// 				$("#"+unique).offset({left: offset.left - subtraioffset});	
// 				//$(".dropup").get(index).style.setProperty("left", offset - subtraioffset);	
// 				$(".dropup-content"+id).get(index).style.setProperty("background-color", background);					
// 				$(".speech-bubble").get(index).style.setProperty("--color", background);																
// 				$(".speech-bubble").get(index).style.setProperty("--dialogposition", dialogposition);																
// 			}	
// 		}else{
// 			if($(".dropup-content"+id).get(index).style.getPropertyValue("display") != "none"){
// 				$(".dropup-content"+id).get(index).style.setProperty("display", "none");						
// 			}
// 		}
// 	});
// });