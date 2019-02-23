using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace importerWZ.Datasources
{
    class UrzZewDataSource
    {
        private UrzZewDataSource.Nabywca nabywca;
        private FbConnection fbConn;
        private static string connectionString;
        private string _typDok;
        private string idKontrahNab;

        public UrzZewDataSource(string typDok, UrzZewDataSource.Nabywca nab)
        {
            this._typDok = typDok;
            this.nabywca = nab;

            UrzZewDataSource.connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();
            if (this.nabywca == UrzZewDataSource.Nabywca.LOTOS)
                this.idKontrahNab = ConfigurationManager.AppSettings["IdKontrahNab_LOTOS"].ToString();
            if (this.nabywca == UrzZewDataSource.Nabywca.BP)
                this.idKontrahNab = ConfigurationManager.AppSettings["IdKontrahNab_BP"].ToString();
            this.fbConn = new FbConnection(UrzZewDataSource.connectionString);
        }

        public void Connect()
        {
            this.fbConn.Open();
        }

        public void Disconnect()
        {
            this.fbConn.Close();
        }

        public bool AddDok(DataRow data, out string message)
        {
            message = "";
            FbTransaction fbTrans = this.fbConn.BeginTransaction();
            try
            {
                string str = (string)null;
                if (this.nabywca == UrzZewDataSource.Nabywca.LOTOS)
                    str = this.GetIdKontrahLotos(data["Konto"].ToString().Trim(), fbTrans);
                if (this.nabywca == UrzZewDataSource.Nabywca.BP)
                    str = this.GetIdKontrahBP(data["Konto"].ToString().Trim(), fbTrans);
                //if (string.IsNullOrEmpty(str))
                // {
                //  fbTrans.Rollback();
                // message += "Nie powiązano odbiorcy z pola 'Konto'";
                // return false;
                // }
                FbCommand fbCommand1 = new FbCommand("execute procedure URZZEWNAGL_ADD_3 (@AID_URZZEWSKAD, @AKODURZ, @ANRUZYT, @ANRAKW, @AODB_IDKONTRAH, @APONIPW, @AJAKINUMERKONTRAH, @AODB_DATA, @AODB_NAZWADOK, @AODB_NRDOK, @AODB_TERMIN, @AODB_PLATNOSC, @AODB_SUMA, @AODB_ILEPOZ, @AODB_GOTOWKA, @AODB_DOT_DATA, @AODB_DOT_NAZWADOK, @AODB_DOT_NRDOK, @AODB_UWAGI, @AODB_DATAOBOW, @AODB_SPOSDOSTAWY, @AODB_CECHA1, @AODB_CECHA2, @AODB_CECHA3, @AODB_CECHA4, @AODB_CECHA5, @AODB_IDNABYWCA, @AODB_IDKONTRAHDOST)");
                fbCommand1.CommandType = CommandType.StoredProcedure;
                fbCommand1.Parameters.Add("@AID_URZZEWSKAD", (object)"2");
                if (this.nabywca == UrzZewDataSource.Nabywca.BP)
                    fbCommand1.Parameters.Add("@AKODURZ", (object)"B2B");
                else
                    fbCommand1.Parameters.Add("@AKODURZ", (object)"B2B");
                fbCommand1.Parameters.Add("@ANRUZYT", (object)"(INT)");
                fbCommand1.Parameters.Add("@ANRAKW", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_IDKONTRAH", (object)this.idKontrahNab);
                fbCommand1.Parameters.Add("@APONIPW", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AJAKINUMERKONTRAH", (object)"2");
                fbCommand1.Parameters.Add("@AODB_DATA", (object)Convert.ToDateTime(data["Data dostawy"].ToString().Trim()).ToString("yyyy-MM-dd"));
                fbCommand1.Parameters.Add("@AODB_NAZWADOK", (object)this._typDok);
                fbCommand1.Parameters.Add("@AODB_NRDOK", (object)data["Opis"].ToString().Trim());
                fbCommand1.Parameters.Add("@AODB_TERMIN", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_PLATNOSC", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_SUMA", (object)"0");
                fbCommand1.Parameters.Add("@AODB_ILEPOZ", (object)"0");
                fbCommand1.Parameters.Add("@AODB_GOTOWKA", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_DOT_DATA", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_DOT_NAZWADOK", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_DOT_NRDOK", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_UWAGI", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_DATAOBOW", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_SPOSDOSTAWY", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_CECHA1", (object)data["Konto"].ToString().Trim());
                fbCommand1.Parameters.Add("@AODB_CECHA2", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_CECHA3", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_CECHA4", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_CECHA5", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AODB_IDNABYWCA", (object)str);
                fbCommand1.Parameters.Add("@AODB_IDKONTRAHDOST", (object)DBNull.Value);
                fbCommand1.Parameters.Add("@AID_URZZEWNAGL", FbDbType.Integer);
                fbCommand1.Parameters["@AID_URZZEWNAGL"].Direction = ParameterDirection.Output;
                fbCommand1.Parameters.Add("@BLAD", FbDbType.Integer);
                fbCommand1.Parameters["@BLAD"].Direction = ParameterDirection.Output;
                fbCommand1.Connection = this.fbConn;
                fbCommand1.Transaction = fbTrans;
                fbCommand1.ExecuteNonQuery();

                if (fbCommand1.Parameters["@BLAD"].Value != DBNull.Value && (uint)Convert.ToInt32(fbCommand1.Parameters["@BLAD"].Value) > 0U)
                {
                    fbTrans.Rollback();
                    message = message + "URZZEWNAGL_ADD " + fbCommand1.Parameters["@BLAD"].Value;
                    return false;
                }
                MessageBox.Show(fbCommand1.Parameters["@AID_URZZEWNAGL"].Value.ToString());

                FbCommand fbCommand2 = new FbCommand("execute procedure URZZEWPOZ_ADD_9 (@AID_URZZEWNAGL, @AKODTOW, @AILOSC, @ACENA, @ACENAUZG, @APROCBONIF, @AODB_UWAGI, @AODB_CECHA1, @AODB_CECHA2, @AODB_CECHA3, @AODB_CENA_BRUTTO, @AODB_DOT_DATA, @AODB_DOT_NAZWADOK, @AODB_DOT_NRDOK, @AODB_DOT_LP, @AODB_MAG_OZNNRWYDR, @AODB_DATA_WAZNOSCI, @AODB_RODZAJ_REZERWACJI, @AODB_ILOSC_REZERWACJI, @AODB_ID_DOSTAWA_REZ, @Aodb_PROCCLA, @ACENA_PRZEPISZ, @AODB_DOT_LPDOD, @AODB_NRDOSTAWY)");
                fbCommand2.CommandType = CommandType.StoredProcedure;
                fbCommand2.Parameters.Add("@AID_URZZEWNAGL", (object)fbCommand1.Parameters["@AID_URZZEWNAGL"].Value.ToString());
                fbCommand2.Parameters.Add("@AKODTOW", (object)data["Indeks"].ToString().Trim());
                if (this.nabywca == UrzZewDataSource.Nabywca.BP)
                    fbCommand2.Parameters.Add("@AILOSC", (object)data["ilość w KG"].ToString().Replace(',', '.').Trim());
                else
                    fbCommand2.Parameters.Add("@AILOSC", (object)data["ilość w L"].ToString().Replace(',', '.').Trim());
                fbCommand2.Parameters.Add("@ACENA", (object)data["Cena Jednostkowa"].ToString().Replace(',', '.').Trim());
                fbCommand2.Parameters.Add("@ACENAUZG", (object)"1");
                fbCommand2.Parameters.Add("@APROCBONIF", (object)"0");
                fbCommand2.Parameters.Add("@AODB_UWAGI", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_CECHA1", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_CECHA2", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_CECHA3", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_CENA_BRUTTO", (object)"0");
                fbCommand2.Parameters.Add("@AODB_DOT_DATA", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_DOT_NAZWADOK", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_DOT_NRDOK", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_DOT_LP", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_MAG_OZNNRWYDR", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_MAG_OZNNRWYDR", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_DATA_WAZNOSCI", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_RODZAJ_REZERWACJI", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_ILOSC_REZERWACJI", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_ID_DOSTAWA_REZ", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@Aodb_PROCCLA", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@ACENA_PRZEPISZ", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_DOT_LPDOD", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AODB_NRDOSTAWY", (object)DBNull.Value);
                fbCommand2.Parameters.Add("@AID_URZZEWPOZ", FbDbType.Integer);
                fbCommand2.Parameters["@AID_URZZEWPOZ"].Direction = ParameterDirection.Output;
                fbCommand2.Parameters.Add("@BLAD", FbDbType.Integer);
                fbCommand2.Parameters["@BLAD"].Direction = ParameterDirection.Output;
                fbCommand2.Connection = this.fbConn;
                fbCommand2.Transaction = fbTrans;

                fbCommand2.ExecuteNonQuery();
                //String id_urzzewnagl = (String)fbCommand1.Parameters["@AID_URZZEWNAGL"].Value;
                // MessageBox.Show(id_urzzewnagl);
                new FbCommand("UPDATE urzzewnagl set status = 0 WHERE id_urzzewnagl=" + fbCommand1.Parameters["@AID_URZZEWNAGL"].Value, this.fbConn)
                {
                    Transaction = fbTrans
                }.ExecuteNonQuery();
                if (fbCommand2.Parameters["@BLAD"].Value != DBNull.Value && (uint)Convert.ToInt32(fbCommand2.Parameters["@BLAD"].Value) > 0U)
                {
                    fbTrans.Rollback();
                    message = message + "URZZEWPOZ_ADD " + fbCommand2.Parameters["@BLAD"].Value;
                    return false;
                }
            }
            catch (Exception ex)
            {
                if (fbTrans != null)
                    fbTrans.Rollback();
                message = message + "UrzZewDataSource() " + ex.Message;
                return false;
            }
            fbTrans.Commit();
            return true;
        }

        private string GetIdKontrahLotos(string konto, FbTransaction fbTrans)
        {
            FbCommand fbCommand = new FbCommand();
            MessageBox.Show(konto);
            fbCommand.CommandText = "SELECT FIRST 1 id_kontrah FROM wystcechykontrah WHERE id_cecha = 10088 AND wartosc = '" + konto + "'";
            fbCommand.Connection = this.fbConn;
            fbCommand.Transaction = fbTrans;
            return fbCommand.ExecuteScalar()?.ToString();
        }

        private string GetIdKontrahBP(string konto, FbTransaction fbTrans)
        {
            FbCommand fbCommand = new FbCommand();
            fbCommand.CommandText = "SELECT FIRST 1 id_kontrah FROM wystcechykontrah WHERE id_cecha = 10124 AND wartosc = '" + konto + "'";
            fbCommand.Connection = this.fbConn;
            fbCommand.Transaction = fbTrans;
            return fbCommand.ExecuteScalar()?.ToString();
        }

        public enum Nabywca
        {
            LOTOS,
            BP,
        }
    }
}
