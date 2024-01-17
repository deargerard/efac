<?php
session_start();
include_once '../call/cone.php';
include_once '../cons.php';
include_once '../call/func.php';
$clientep=iseguro($cone,$_GET['cl']);
$iniped=fmysql(iseguro($cone,$_GET['fi']));
$finped=fmysql(iseguro($cone,$_GET['ff']));
$estado=iseguro($cone,$_GET['est']);
if(isset($clientep) && !empty($clientep) && isset($iniped) && !empty($iniped) && isset($finped) && !empty($finped) && isset($estado) && !empty($estado)){
      $fecha = @date("d-m-Y");

      //Inicio de la instancia para la exportación en Excel
      //header('Content-type: application/vnd.ms-excel');
      header("Content-Type:   application/vnd.ms-excel; charset=utf-8");
      header("Content-Disposition: attachment; filename=ventas_$iniped-al-$finped.xls");
      header("Pragma: no-cache");
      header("Expires: 0");
?>
      <table border=1>
<?php
          if ($clientep=="t") {
              $wu="";
            }else {
              $wu= "AND p.idpersona=$clientep";
            }
            $c="SELECT v.idventa, v.fecha, dv.numero, td.s_tipcom, l.seriecom, p.tipodocumento, p.numerodoc, p.nombre, v.estado FROM venta v INNER JOIN docventa dv ON v.idventa=dv.idventa INNER JOIN tipodoc td ON dv.idtipodoc=td.idtipodoc INNER JOIN persona p ON v.cliente=p.idpersona INNER JOIN local l ON v.idlocal=l.idlocal WHERE (v.estado=2 OR v.estado=3) AND (date_format(v.fecha, '%Y-%m-%d') BETWEEN '$iniped' AND '$finped') $wu";
            //$c="SELECT v.idventa, u.usuario as cajero, us.usuario as despachador, l.nombre AS tienda, v.fecha, p.idpersona, p.nombre, v.descuento, tp.tipopag, td.tipodoc, LPAD(dv.numero, 8,'0') AS numdoc, l.seriecom FROM venta v INNER JOIN persona p ON v.cliente=p.idpersona INNER JOIN local l ON v.idlocal=l.idlocal INNER JOIN caja c ON c.idcaja=v.idcaja INNER JOIN usuario u ON c.idusuario = u.idusuario INNER JOIN usuario us ON v.idusuario=us.idusuario INNER JOIN tipopag tp ON v.idtipopag= tp.idtipopag INNER JOIN docventa dv ON v.idventa=dv.idventa INNER JOIN tipodoc td ON dv.idtipodoc=td.idtipodoc WHERE v.estado=$estado AND (date_format(v.fecha, '%Y-%m-%d') BETWEEN '$iniped' AND '$finped') $wu";
            $cp=mysqli_query($cone,$c);
            if (mysqli_num_rows($cp)>0) {
?>
                <tr>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ffechadoc D</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ffechaven D</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ccoddoc C(2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cserie C(20)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cnumero C(20)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ccodenti C(11)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cdesenti C(100)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ctipdoc C(1)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ccodruc C(15)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">crazsoc C(100)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nbase2 N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nbase1 N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nexo N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nina N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nisc N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nigv1 N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nicbpers N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nbase3 N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ntots N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ntc N(10,6)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">freffec D</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">crefdoc C(2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">crefser C(6)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">crefnum C(13)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cmreg C(1)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ndolar N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ffechaven2 D</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ccond C(3)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cccodcos C(9)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ccodcos2 C(9)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cctabase C(20)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cctaicbper C(20)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cctaotrib C(20)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cctatot C(20)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nresp N(1)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nporre N(5,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nimpres N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cserre C(6)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cnumre C(13)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ffecre D</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ccodpresu C(10)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nigv N(5,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cglosa C(80)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">ccodpago C(3)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nperdenre N(1)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">nbaseres N(15,2)</font></td>
                  <td bgcolor= "#777777"><font color="#ffffff" size="2">cctaperc C(20)</font></td>
                </tr>
                <?php
                while($rcp=mysqli_fetch_assoc($cp)){
                    $idped=$rcp['idventa'];

                    switch ($rcp['s_tipcom']) {
                      case '01':
                        $pco="F";
                        break;
                      case '03':
                        $pco="B";
                          break;
                      default:
                        $pco="T";
                        break;
                    }

                    $cdv=mysqli_query($cone,"SELECT SUM(subtotal) st FROM detventa WHERE idventa=$idped;");
                      if($rdv=mysqli_fetch_assoc($cdv)){
                        $pv=round($rdv['st'],2);
                        $pc=round($pv/1.18,2);
                        $igv=round($pv-$pc,2);

                        if($rcp['estado']==3){
                          switch ($rcp['tipodocumento']) {
                            case '1':
                              $dd='=TEXTO(0;"00000000")';
                              break;
                            case '6':
                              $dd='=TEXTO(0;"00000000000")';
                                break;
                            default:
                              $dd="0";
                              break;
                          }
                          $ntd=0;
                          $ndd='=TEXTO(0;"00000000")';
                          $nnom="ANULADO";

                          $pv=0;
                          $pc=0;
                          $igv=0;
                        }else{
                          if($rcp['nombre']=="SIN NOMBRE"){
                            $dd='=TEXTO(0;"00000000")';
                            $ntd=$rcp['tipodocumento'];
                            $ndd="999999999";
                            $nnom="BOLETA DE VENTA";
                          }else{
                            $dd=$rcp['numerodoc'];
                            $ntd=$rcp['tipodocumento'];
                            $ndd=$rcp['numerodoc'];
                            $nnom=$rcp['nombre'];
                          }
                        }
                ?>
                <tr>
                  <td><font size="2"><?php echo date('d/m/Y', strtotime($rcp['fecha'])); ?></font></td>
                  <td><font size="2"><?php echo date('d/m/Y', strtotime($rcp['fecha'])); ?></font></td>
                  <td><font size="2"><?php echo '=TEXTO('.$rcp['s_tipcom'].';"00")'; ?></font></td>
                  <td><font size="2"><?php echo $pco.$rcp['seriecom']; ?></font></td>
                  <td><font size="2"><?php echo '=TEXTO('.$rcp['numero'].';"0000000000")'; ?></font></td>
                  <td><font size="2"><?php echo $rcp['tipodocumento']; ?></font></td>
                  <td><font size="2"><?php echo $dd; ?></font></td>
                  <td><font size="2"><?php echo $ntd; ?></font></td>
                  <td><font size="2"><?php echo $ndd ?></font></td>
                  <td><font size="2"><?php echo $nnom; ?></font></td>
                  <td><font size="2"></td>
                  <td><font size="2"><?php echo $pc; ?></font></td>
                  <td><font size="2">0.00</font></td>
                  <td><font size="2">0.00</font></td>
                  <td><font size="2">0.00</font></td>
                  <td><font size="2"><?php echo $igv; ?></font></td>
                  <td><font size="2">0.00</font></td>
                  <td><font size="2">0.00</font></td>
                  <td><font size="2"><?php echo $pv; ?></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2">S</font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"><?php echo date('d/m/Y', strtotime($rcp['fecha'])); ?></font></td>
                  <td><font size="2">CON</font></td>
                  <td><font size="2">020101</font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2">70121</font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2">1011</font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2">18.00</font></td>
                  <td><font size="2">VENTA DE MERCADERIAS LOCAL - EFECTIVO</font></td>
                  <td><font size="2">=TEXTO(9;"000")</font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                  <td><font size="2"></font></td>
                </tr>
<?php
                      }
                }
?>
          </table>
<?php
                }else{
?>                 <tr>
                    <td colspan="20">NO EXISTEN VENTAS QUE MOSTRAR</td>
                  </tr>
<?php
                }
}else{
		echo mensajeda("Error: Debe seleccionar al menos un valor en cada campo");
	}
  mysqli_close($cone);
?>
