<!DOCTYPE html>
<html lang="en">

<body>
  <div class="container-scroller">
    <div class="container-fluid page-body-wrapper full-page-wrapper">
      <div class="main-panel" style="background-image: url(/images/fundo.jpg);background-position: center; background-repeat: no-repeat; background-size: cover;">
        <div class="content-wrapper d-flex align-items-center auth px-0">
          <div class="row w-100 mx-0">
            <div class="col-lg-4 mx-auto">
              <div class="auth-form-light text-left py-5 px-4 px-sm-5">
                <div class="brand-logo" style="display:clock;margin-left:auto;margin-right: auto;">
                  <img src="../images/logo_extenso_sem_fundo.png" style = "display: block; margin-left: auto; margin-right: auto;width:100%" alt="logo">
                </div>
                <div style="text-align: center;">
                  <h4>Esqueceu sua senha?</h4>
                  <h6 class="font-weight-light">Digite o email cadastrado abaixo e enviaremos instruções para recuperação</h6>
                </div>
                <form class="pt-3" method="POST">
                  <div class="form-group">
                    <input type="email" class="form-control form-control-lg" id="exampleInputEmail1" name="email" placeholder="Email" required="required">
                  </div>
                  <div class="mt-3">
                    <button class="btn btn-block btn-dark btn-lg font-weight-medium auth-form-btn" formaction="?pg=emailRecover">Recuperar</button>
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

