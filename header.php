<?php	
	require_once("./config/config.php");
    $dataReceived = (file_get_contents("php://input"));
	if($dataReceived!="");
		$session = json_decode($dataReceived);

	$host = "https://".$_SERVER['HTTP_HOST'];
?>
<html>
	<body>
		<div class="horizontal-menu">
		<nav class="navbar top-navbar col-lg-12 col-12 p-0">
			<div class="container-fluid">
			<div class="navbar-menu-wrapper d-flex align-items-center justify-content-between">
				<ul class="navbar-nav navbar-nav-left">
				<div class="text-center navbar-brand-wrapper d-flex align-items-center justify-content-center">
					<a class="navbar-brand brand-logo" ><img src=<?php echo $host."/images/logo_redonda.png";?> alt="logo"/></a>
				</div>
				</ul>
				<ul class="navbar-nav" style = "margin-left:170;">
				<div class="text-center navbar-brand-wrapper d-flex align-items-center justify-content-center">
					<span class="nav-profile-name">
						<a class="navbar-brand brand-logo" style="width: %;" ><img src=<?php echo $host."/images/logo_letra.png";?> alt="logo"/></a>
					</span>
				</div>
				</ul>
				<ul class="navbar-nav navbar-nav-right">
					<li class="nav-item nav-profile dropdown">
					<a class="nav-link dropdown-toggle" href="#" data-toggle="dropdown" id="profileDropdown">
						<span class="nav-profile-name"> <?php if(isset($session)) echo $session->nome//colocar um getUser aq depois?> </span>
						<span class="online-status"></span>
						<?php 
							if(isset($session->file)) 
								if($session->file != "")
									echo '<img src="'.'data:image/png;base64,'. $session->file .'" width="30" height="30" id="imgSession" alt="profile"/>';
								else
									echo '<img src="'.$host.'/images/user_default.png" width="25" height="25" alt="profile" style="margin-top:-2px;margin-right:4px"/>';
							else
								echo '<img src="'.$host.'/images/user_default.png" width="25" height="25" alt="profile" style="margin-top:-2px;margin-right:4px"/>';
						 ?> 
					</a>
					<div class="dropdown-menu dropdown-menu-right navbar-dropdown" aria-labelledby="profileDropdown" >
						<a class="dropdown-item" href="./index.php?pg=perfil"  isGlobalLink="isGlobalLink">
							<i class="mdi mdi-account-outline text-primary"></i>
							Perfil
						</a>
						<a class="dropdown-item" href="./index.php?pg=doLogout" isGlobalLink="isGlobalLink">
							<i class="mdi mdi-logout text-primary"></i>
							Sair
						</a>
					</div>
					</li>
				</ul>
				<button class="navbar-toggler navbar-toggler-right d-lg-none align-self-center" type="button" data-toggle="horizontal-menu-toggle">
				<span class="mdi mdi-menu"></span>
				</button>
			</div>
			</div>
		</nav>
		<nav class="bottom-navbar">
			<div class="container">
			<div class="col-md-12 page-alerts"></div>
				<ul class="nav page-navigation">
                <?php 
					$user = new Usuarioo($session->id);
					$tipoAcesso = $user->idTipoAcesso_id;
					switch($tipoAcesso){
						case 1:
							$menu = new ItemMenu("Equipamentos", "all-inclusive","*equipamento");
							$menu->show();
							$menu = new ItemMenu("Modalidades", "delta","*modalidades");
							$menu->show();
							$menu = new ItemMenu("Clientes", "beta","*clientes");
							$menu->show();
							$menu = new ItemMenu("Personais", "lambda","*personal");
							$menu->show();
							$menu = new ItemMenu("Funcionários", "omega","*funcionario");
							$menu->show();
							$menu = new ItemMenu("Usuários", "sigma","*usuarios");
							$menu->show();
						break;
						case 2: // Aluno
							$menu = new ItemMenu("Modalidades", "beta","*modalidades");
							$menu->show();
							$menu = new ItemMenu("Treinos", "delta","*treinos");
							$menu->show();
							$menu = new ItemMenu("Personais", "lambda","*personal");
							$menu->show();
							$menu = new ItemMenu("Equipamentos", "all-inclusive","*equipamento");
							$menu->show();
						break;
						case 3:
							$menu = new ItemMenu("Equipamentos", "all-inclusive","*equipamento");
							$menu->show();
							$menu = new ItemMenu("Modalidades", "delta","*modalidades");
							$menu->show();
							$menu = new ItemMenu("Treinos", "lambda","*treinos");
							$menu->show();
							$menu = new ItemMenu("Clientes", "beta","*clientes");
							$menu->show();
							$menu = new ItemMenu("Funcionários", "omega","*funcionario");
							$menu->show();
						break;
					}
                ?>
				</ul>
			</div>
		</nav>
		</div>
	<div class="container-fluid page-body-wrapper">
			<div class="main-panel">
				<div class="content-wrapper">
	</body>


</html>