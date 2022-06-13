<!DOCTYPE html>
<html lang="en">

<body >
  <div class="container-scroller" >
    <div class="container-fluid page-body-wrapper full-page-wrapper">
      <div class="main-panel" style="background-image: url(/images/fundo.jpg);background-position: center; background-repeat: no-repeat; background-size: cover;">
        <div class="content-wrapper d-flex align-items-center auth px-0">
          <div class="row w-100 mx-0">
            <div class="col-lg-4 mx-auto">
              <div class="auth-form-light text-left py-5 px-4 px-sm-5">
                <div class="brand-logo">
                  <img src="../images/logo_extenso_sem_fundo.png" style = "display: block; margin-left: auto; margin-right: auto; width: 100%;" alt="logo">
                </div>
                <h4>Bem vindo a Sistema de Academia</h4>
                <h6 class="font-weight-light">Entre com seus dados para continuar</h6>
                <form id="form-login" class="pt-3" method="POST">
                  <div class="form-group">
                    <input type="text" required class="form-control form-control-lg" id="exampleInputEmail1" name="login" placeholder="Usuário">
                  </div>
                  <div class="form-group">
                    <input type="password" required class="form-control form-control-lg" id="exampleInputPassword1" name="password" placeholder="Senha">
                  </div>
                  <div class="mt-3">
                    <button input-type=submit formaction="?pg=doLogin" class="btn btn-block btn-dark btn-lg font-weight-medium auth-form-btn">ENTRAR</button>
                  </div>
                  <div class="my-2 d-flex justify-content-between align-items-center">
                    <div class="form-check">
                      <label class="form-check-label text-muted">
                        <input type="checkbox" class="form-check-input">
                        Lembrar-se de mim
                      </label>
                    </div>
                    <a href="?pg=doRecover" class="auth-link text-black">Esqueceu a senha?</a>
                  </div>
                  <!-- <div class="mb-2">
                    <button type="button" class="btn btn-block btn-facebook auth-form-btn">
                      <i class="mdi mdi-facebook mr-2"></i>Connect using facebook
                    </button>
                  </div> -->
                  <div class="text-center mt-4 font-weight-light">
                    Não tem uma conta ainda? <a href="?pg=register" class="text-primary">Criar</a>
                  </div>
                </form>
              </div>
            </div>
          </div>
        </div>
      </div>
      <!-- content-wrapper ends -->
    </div>
    <!-- page-body-wrapper ends -->
  </div>
</body>
<script src="../Framework/framework.js"></script>
<script src="./javascript/login.js"></script>
</html>
