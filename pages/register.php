<!DOCTYPE html>
<html lang="en">

<body>
  <div class="container-scroller">
    <div class="container-fluid page-body-wrapper full-page-wrapper">
      <div class="main-panel" style="background-image: url(/images/fundo.jpg);background-position: center; background-repeat: no-repeat; background-size: cover;">
        <div class="content-wrapper d-flex align-items-center auth px-0">
          <div class="row w-100 mx-0">
            <div class="col-lg-8 mx-auto">
              <div class="auth-form-light text-left py-4 px-3 px-sm-4">
                <div class="brand-logo">
                  <img src="../images/logo_extenso_sem_fundo.png" style = "display: block; margin-left: auto; margin-right: auto; width: 50%;" alt="logo">
                </div>
                 <div style="text-align: center;">
                  <h4>Novo por aqui?</h4>
                  <h6 class="font-weight-light">Criar sua conta é fácil, só leva alguns minutos</h6>
                </div>
                <form class="pt-2 row" method="POST">
                  <div class="form-group col-lg-12">
                    <input type="text" class="form-control form-control-lg" id="nome" name="nome" placeholder="Nome Completo">
                  </div>
                  <div class="form-group col-lg-7">
                    <input type="email" class="form-control form-control-lg" id="exampleInputEmail1" name="email" placeholder="Email">
                  </div>
                  <div class="form-group col-lg-5">
                    <input type="text" class="form-control form-control-lg" id="exampleInputcelular1" name="telefone" placeholder="Telefone">
                  </div>
                  <div class="form-group col-lg-7">
                    <input type="text" class="form-control form-control-lg" id="exampleInputcpf1" name="CPF" placeholder="CPF">
                  </div>
                  <div class="form-group col-lg-5">
                    <input type="date" class="form-control form-control-lg" id="exampleInputidade1" name="dataNascimento" placeholder="Idade">
			  	          <label class="help-block"> Data de Nascimento</label>        
                  </div>
                  <div class="form-group col-lg-12">
                    <input type="text" class="form-control form-control-lg" id="exampleInputenderco1" name="endereco" placeholder="Endereço por Extenso">
                  </div>
                  <div class="form-group col-lg-6">
                    <input type="text" class="form-control form-control-lg" id="exampleInputUsername1" name="login" placeholder="Usuário">
                  </div>
                  <div class="form-group col-lg-6">
                    <input type="password" class="form-control form-control-lg" id="exampleInputPassword1" name="password" placeholder="Password">
                  </div>


                  <div class="mt-3">
                    <button class="btn btn-block btn-dark btn-lg font-weight-medium auth-form-btn" formaction="?pg=doRegister">Cadastrar</button>
                  </div>
                  <div class="text-center mt-4 font-weight-light">
                    Já tem uma conta? <a href="?pg=login" class="text-primary">Entre</a>
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

</html>

