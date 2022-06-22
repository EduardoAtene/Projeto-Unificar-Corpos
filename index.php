<?php 
	require_once("./config/config.php");
	if(session_id() == '')
		session_start(); 
	ob_start();
	//Identifica pro sistema que tipo de página estamos abrindo ("redirect" significa que um modulo está sendo acessado)
	$pageHub = Utility::getVariable('pg', INPUT_GET);
	//Identifica QUAl módulo está sendo acessado, se for o caso
	$System = Utility::getVariable("system", INPUT_GET);
	//Path padrão do sistema
	$path = $_SERVER['DOCUMENT_ROOT']."pages/";
	$pageExtension = ".php";
	$host = "https://".$_SERVER['HTTP_HOST'];
	?>
	<html>
		<head>
			<!-- Required meta tags -->
			<meta charset="utf-8">
			<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>

			<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
			<title>Projeto Unificar Corpo</title>
			<!-- base:css -->
			<link rel="stylesheet" href=<?php echo $host."/vendors/mdi/css/materialdesignicons.min.css";?>>
			<link rel="stylesheet" href=<?php echo $host."/vendors/base/vendor.bundle.base.css";?>>
			<link rel="stylesheet" href="//cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css">
			<link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
			<!-- CSS only -->
			<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous">
			<link rel="stylesheet" href="https://code.jquery.com/ui/1.9.0/themes/base/jquery-ui.css" />
			<!-- endinject -->
			<!-- plugin css for this page -->
			<!-- End plugin css for this page -->
			<!-- inject:css -->
			<link rel="stylesheet" href=<?php echo $host."/css/style.css";?>>
			<!-- endinject -->
			 <link rel="shortcut icon" href="images/logo_redonda.ico" />
		</head>
		<!-- base:js -> carrega o jquery -->
		<script src=<?php echo $host."/vendors/base/vendor.bundle.base.js";?>></script>
		<!-- endinject -->
		<!-- inject:js -->
		<script src=<?php echo $host."/js/template.js";?>></script>
		<script src=<?php echo $host."/javascript/index.js";?>></script>
		<script src="https://code.jquery.com/jquery-1.8.2.js"></script>
		<script src="https://code.jquery.com/ui/1.9.0/jquery-ui.js"></script>
		<script>
		$(function() {
			$( 'input[input-type="date"]' ).datepicker({
															dateFormat: 'dd/mm/yy',
															dayNames: ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo'],
															dayNamesMin: ['D','S','T','Q','Q','S','S','D'],
															dayNamesShort: ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb','Dom'],
															monthNames: ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'],
															monthNamesShort: ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
														});
		});
		</script>
		<!-- <script src="javascript/jquery-1.11.2.js"></script>
		<script src="javascript/jquery-ui.min.js"></script> -->
		<!-- endinject -->
		<script src=<?php echo $host."/vendors/justgage/raphael-2.1.4.min.js";?>></script>
		<script src=<?php echo $host."/vendors/justgage/justgage.js";?>></script>
		<script src=<?php echo $host."/vendors/progressbar.js/progressbar.min.js";?>></script>
		<script src=<?php echo $host."/Framework/framework.js";?>></script>
		<script src="//cdn.datatables.net/1.10.25/js/jquery.dataTables.min.js"></script>
		<script src="//cdn.jsdelivr.net/npm/sweetalert2@11"></script>
		<script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
		<script src="https://cdn.jsdelivr.net/npm/cleave.js@1.6.0/dist/cleave.min.js"></script> <!-- formatar inputs -->
		<!-- JavaScript Bundle with Popper -->
		<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
		<script src=<?php echo $host."/vendors/chart.js/Chart.min.js";?>></script>
		<script src=<?php echo $host."/js/chart.js";?>></script>
		<script type="text/javascript" src="js/jquery.maskedinput-1.1.4.pack.js"/></script>
	</html>
	<?php

	$loginFailWarning = '<div class="pro-banner" id="pro-banner">
							<div class="card pro-banner-bg border-0 rounded-0">
								<div class="card-body py-3 px-4 d-flex align-items-center justify-content-between flex-wrap">
									<p class="mb-0 text-white font-weight-medium mb-2 mb-lg-0 mb-xl-0">Seu login e/ou senha não estão de acordo com o cadastrado. Por favor tente novamente com as credenciais corretas.</p>
									<div class="d-flex">
										<button id="bannerClose" class="btn border-0 p-0">
											<i class="mdi mdi-close text-white"></i>
										</button>
									</div>
								</div>
							</div>
						</div>';
	$AlreadyHasUserWarning = '<div class="pro-banner" id="pro-banner">
								<div class="card pro-banner-bg border-0 rounded-0">
									<div class="card-body py-3 px-4 d-flex align-items-center justify-content-between flex-wrap">
										<p class="mb-0 text-white font-weight-medium mb-2 mb-lg-0 mb-xl-0">Desculpe, mas o login informado já consta em nossa base de dados. Por favor escolha outro ou entre utilizando a senha correspondente.</p>
										<div class="d-flex">
											<button id="bannerClose" class="btn border-0 p-0">
												<i class="mdi mdi-close text-white"></i>
											</button>
										</div>
									</div>
								</div>
							</div>';
	$UserCreated = '<div class="pro-banner" id="pro-banner">
							<div class="card pro-banner-bg border-0 rounded-0">
								<div class="card-body py-3 px-4 d-flex align-items-center justify-content-between flex-wrap">
									<p class="mb-0 text-white font-weight-medium mb-2 mb-lg-0 mb-xl-0">Usuário criado com sucesso. Entre com as credenciais informadas</p>
									<div class="d-flex">
										<button id="bannerClose" class="btn border-0 p-0">
											<i class="mdi mdi-close text-white"></i>
										</button>
									</div>
								</div>
							</div>
						</div>';

	if(!Session::hasSession()){

		if($pageHub!=""){
			if($pageHub=="doRecover" || $pageHub=="emailRecover" || $pageHub=="planos" || $pageHub=="testEmail" || $pageHub=="register" || $pageHub=="doRegister" || $pageHub=="recuperarSenha" || $pageHub=="logout" || $pageHub=="login" || $pageHub=="doLogin" || $pageHub=="curriculo" || $pageHub=="talentos" || $pageHub=="correios"){
				include($path.$pageHub.$pageExtension);
			}else{
				if(array_key_exists("statusLogin",$_SESSION)){
					if($_SESSION["statusLogin"]=="fail"){
						echo $loginFailWarning;
						unset($_SESSION['statusLogin']);
					}
					else if($_SESSION["statusLogin"]=="alreadyHasUser"){
						echo $AlreadyHasUserWarning;
						unset($_SESSION['statusLogin']);
					}
					else if($_SESSION["statusLogin"]=="UserCreated"){
						echo $UserCreated;
						unset($_SESSION['statusLogin']);
					}
				}
				include($path."login.php");
			}			
		}else{
			if(array_key_exists("statusLogin",$_SESSION)){
				if($_SESSION["statusLogin"]=="fail"){
					echo $loginFailWarning;
					unset($_SESSION['statusLogin']);
				}
				else if($_SESSION["statusLogin"]=="alreadyHasUser"){
					echo $AlreadyHasUserWarning;
					unset($_SESSION['statusLogin']);
				}
				else if($_SESSION["statusLogin"]=="UserCreated"){
					echo $UserCreated;
					unset($_SESSION['statusLogin']);
				}
			}
			include($path."login.php");
		}
	}
	else{
		switch($pageHub){

			case "doLogout":
				include($path.$pageHub.$pageExtension);
			break;
			case null:
				$webPage = Utility::get_web_page("https://".$_SERVER['HTTP_HOST']."/header.php",json_encode($_SESSION));
				echo $webPage["content"];
			break;
			case "analises":
				$webPage = Utility::get_web_page("https://".$_SERVER['HTTP_HOST']."/header.php",json_encode($_SESSION));
				echo $webPage["content"];
				include($path.$pageHub.$pageExtension);
				//$webPage = Utility::get_web_page(",json_encode($_SESSION));
				//var_dump($webPage);
			break;
			default:
				$webPage = Utility::get_web_page("https://".$_SERVER['HTTP_HOST']."/header.php",json_encode($_SESSION));
				echo $webPage["content"];

				$webPage = Utility::get_web_page("https://".$_SERVER['HTTP_HOST']."/pages/".$pageHub.$pageExtension,json_encode($_SESSION));
				echo $webPage["content"];
			break;
		}
	}
	exit();

	/**
 * Get a web file (HTML, XHTML, XML, image, etc.) from a URL.  Return an
 * array containing the HTTP server response header fields and content.
 */
function get_web_page( $url )
{
	$cookieFile = "cookies.txt";
	if(!file_exists($cookieFile)) {
		$fh = fopen($cookieFile, "w");
		fwrite($fh, "");
		fclose($fh);
	}
    $options = array(
        CURLOPT_RETURNTRANSFER => true,     // return web page
        CURLOPT_HEADER         => false,    // don't return headers
        CURLOPT_FOLLOWLOCATION => true,     // follow redirects
        CURLOPT_ENCODING       => "",       // handle all encodings
        CURLOPT_USERAGENT      => "spider", // who am i
        CURLOPT_AUTOREFERER    => true,     // set referer on redirect
        CURLOPT_CONNECTTIMEOUT => 120,      // timeout on connect
        CURLOPT_TIMEOUT        => 120,      // timeout on response
        CURLOPT_MAXREDIRS      => 10,       // stop after 10 redirects
        CURLOPT_SSL_VERIFYPEER => false,     // Disabled SSL Cert checks
    );
	//curl_setopt($ch, CURLOPT_USERPWD, $username . ":" . $password);  

    $ch      = curl_init( $url );
    curl_setopt_array( $ch, $options );
    $content = curl_exec( $ch );
    $err     = curl_errno( $ch );
    $errmsg  = curl_error( $ch );
    $header  = curl_getinfo( $ch );
    curl_close( $ch );

    $header['errno']   = $err;
    $header['errmsg']  = $errmsg;
    $header['content'] = $content;
    return $header;
}
?>
