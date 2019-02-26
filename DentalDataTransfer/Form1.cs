using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DentalDataTransfer
{
    public partial class Form1 : Form
    {
        private DataTable _tblCustomer;
        private DataTable _tblProcess;
        private DataTable _tblAppointment;

        private DataTable _tblPatients;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _tblCustomer = GetCustomer();
            _tblProcess = GetProcess();
            _tblAppointment = GetAppointment();

            _tblPatients = GetPatients();

            //var list = _tblPatients.Rows.Cast<DataRow>().GroupBy(x => new[] { x[1].ToString(), x[2].ToString() }).Select(x => new { Key = x.Key, C = x.Count() }).Where(x=> x.C > 1).ToList();
        }

        private DataTable GetPatients()
        {
            return GetData("database", "SELECT * FROM PATIENT");
        }

        private void AddPatients(DataTable table)
        {
            string query = @"INSERT INTO PATIENT (
    [NAME],
    [SURNAME],
    [ADDRESS],
    [PHONE_NUMBER],
    [BOOK_NAME],
    [BOOK_PAGE_NUMBER],
    [CREATED_DATE],
    [TEETH_MAP]
)
VALUES (
    @p1,
    @p2,
    @p3,
    @p4,
    '',
    0,
    @p5,
    '11111111111111111111111111111111'
)";

            string cs = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = |DataDirectory|\\database.accdb;";

            using (OleDbConnection connection = new OleDbConnection(cs))
            {
                connection.Open();

                foreach (DataRow row in table.Rows)
                {
                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = query;

                        command.Parameters.AddWithValue("@p1", row[1]);
                        command.Parameters.AddWithValue("@p2", row[2]);
                        command.Parameters.AddWithValue("@p3", row[3]);
                        command.Parameters.AddWithValue("@p4", row[4]);
                        command.Parameters.AddWithValue("@p5", DateTime.ParseExact(row[9].ToString(), "dd.MM.yyyy HH:mm", null));

                        command.ExecuteScalar();
                    }
                }
            }
        }

        private void AddTreatments(DataTable table)
        {
            string query = @"INSERT INTO [TREATMENT] (
    [PATIENT_ID],
    [DESCRIPTION],
    [PRICE],
    [PAID],
    [CREATED_DATE]
) VALUES (
    @p1, @p2, @p3, @p4, @p5
)";

            string cs = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = |DataDirectory|\\database.accdb;";

            using (OleDbConnection connection = new OleDbConnection(cs))
            {
                connection.Open();

                foreach (DataRow row in table.Rows)
                {
                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = query;

                        if (row["Kimlik"].ToString() == "8390")
                        {
                        }

                        double miktar = OdemeMiktariBul(row[3].ToString());
                        bool odendiMiMiktarAlanindan = OdemeTuruAyarlaMiktarAlanindan(row[3].ToString());
                        bool odendiMi = OdenmeTuruAyarla(row[4].ToString());

                        string eskiId = row[1].ToString();
                        DataRow cust = _tblCustomer.Rows.Cast<DataRow>().FirstOrDefault(x => x[0].ToString() == eskiId);

                        if (cust == null)
                        {
                            continue;
                        }

                        int eskiIdIndex = _tblCustomer.Rows.IndexOf(cust);
                        //int yeniId = (int)_tblPatients.Rows[eskiIdIndex][0];
                        int yeniId = (int)_tblPatients.Rows.Cast<DataRow>()
                            .Where(x => x[1].ToString() == cust[1].ToString() && x[2].ToString() == cust[2].ToString())
                            .FirstOrDefault()[0];

                        command.Parameters.AddWithValue("@1", yeniId);
                        command.Parameters.AddWithValue("@2", row[2].ToString());
                        command.Parameters.AddWithValue("@3", miktar);
                        command.Parameters.AddWithValue("@4", odendiMiMiktarAlanindan || odendiMi ? miktar : 0);
                        command.Parameters.AddWithValue("@5", DateTime.ParseExact(row[5].ToString(), "dd.MM.yyyy HH:mm", null));

                        int affectedRows = command.ExecuteNonQuery();

                        if (affectedRows != 1)
                        {
                            MessageBox.Show("Error");

                            return;
                        }
                    }
                }
            }
        }

        private double OdemeMiktariBul(string deger)
        {
            deger = deger.Replace("TL", string.Empty).Replace("tl", string.Empty).Replace(",", string.Empty);
            string[] kelimeler = { "TOPLAM", "ALINDI", "alındı", "EVET", "SURİYELİ", "VLDAN", "HNM", "YAPTI", "POS", "YENİLENDİ", "-", "EURO.", "ALDIM" };

            kelimeler.ToList().ForEach(x => deger = deger.Replace(x, string.Empty));

            deger = deger.Trim();

            if (string.IsNullOrWhiteSpace(deger))
            {
                return 0;
            }

            if (deger == "100  100")
                return 100;

            if (deger.Any(x => !char.IsDigit(x)))
            {
                MessageBox.Show(deger);
            }

            return double.Parse(deger);
        }

        private bool OdemeTuruAyarlaMiktarAlanindan(string deger)
        {
            string[] kelimeler = { "ALINDI", "alındı", "EVET", "POS", "ALDIM" };

            return kelimeler.Any(x => deger.Contains(x));
        }

        private bool OdenmeTuruAyarla(string odendi)
        {
            if (string.IsNullOrWhiteSpace(odendi)) return false;

            if (odendi == "ÖDENMEDİ") return false;

            if (odendi == "-") return false;

            return true;
        }

        private void AddAppointments(DataTable table)
        {
            string query = @"INSERT INTO APPOINTMENT (
    [DATE],
    [ROW_NUMBER],
    [NAME],
    [DESCRIPTION],
    [PHONE_NUMBER],
    [COLOR_NAME],
    [COLOR_DESCRIPTION],
    [COLOR_PHONE_NUMBER]
) VALUES (
    @p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8
)";

            string cs = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = |DataDirectory|\\database.accdb;";

            using (OleDbConnection connection = new OleDbConnection(cs))
            {
                connection.Open();

                foreach (DataRow row in table.Rows)
                {
                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = query;

                        command.Parameters.AddWithValue("@1", DateTime.ParseExact(row[3].ToString(), "dd.MM.yyyy", null));
                        command.Parameters.AddWithValue("@2", int.Parse(row[4].ToString()));
                        command.Parameters.AddWithValue("@3", row[0]);
                        command.Parameters.AddWithValue("@4", row[1]);
                        command.Parameters.AddWithValue("@5", row[2]);
                        command.Parameters.AddWithValue("@6", RenkAyarla(row[5].ToString()));
                        command.Parameters.AddWithValue("@7", RenkAyarla(row[6].ToString()));
                        command.Parameters.AddWithValue("@8", RenkAyarla(row[7].ToString()));

                        int affectedRows = command.ExecuteNonQuery();

                        if (affectedRows != 1)
                        {
                            MessageBox.Show("Error");

                            return;
                        }
                    }
                }
            }
        }

        private int RenkAyarla(string renk)
        {
            if (string.IsNullOrWhiteSpace(renk)) return 0;

            if (renk == "true" || renk == "false") return 0;

            return int.Parse(renk);
        }

        private DataTable GetCustomer()
        {
            // Kimlik, CustomerName, CustomerSurname, CustomerPlace, CustomerPhoneNo, CustomerBookNo, CustomerPageNo, X, X, CustomerRegisterDate
            return GetData("db1", "SELECT * FROM customer WHERE CustomerIsActive = 'true'");
        }

        private DataTable GetProcess()
        {
            // Kimlik, CustomerNo, ProcessInfo, ProcessPrice, ProcessStatus, ProcessDate
            // ProcessPrice: 
            //  if contains "ALINDI"
            return GetData("db2", "SELECT * FROM process");
        }

        private DataTable GetAppointment()
        {
            // NameAndSurname, Info, PhoneNo, AppDate, AppIndex, Color1, Color2, Color3
            //                                                   if not true
            return GetData("db3", "SELECT * FROM appointment");
        }

        private DataTable GetData(string db, string query)
        {
            string cs = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = |DataDirectory|\\{db}.accdb;";

            using (OleDbConnection connection = new OleDbConnection(cs))
            {
                connection.Open();

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    DataTable table = new DataTable();

                    adapter.Fill(table);

                    return table;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //AddPatients(_tblCustomer);
            AddTreatments(_tblProcess);
            //AddAppointments(_tblAppointment);
        }

        #region EXCELDEKİ VERİLERİ ACCESSE AKTARIR

        private void button3_Click(object sender, EventArgs e)
        {
            string cs = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = |DataDirectory|\\excel.xls; Extended Properties=\"Excel 8.0; HDR=NO\"";

            DataTable table = new DataTable();

            using (OleDbConnection connection = new OleDbConnection(cs))
            {
                connection.Open();

                string query = "SELECT * FROM [Sayfa1$]";

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    adapter.Fill(table);
                }
            }

            AddCustomers(table);
        }

        private void AddCustomers(DataTable dt)
        {
            string cs = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = |DataDirectory|\\db1.accdb;";

            string query = @"INSERT INTO [customer] (
    CustomerName,
    CustomerSurname,
    CustomerPlace,
    CustomerPhoneNo,
    CustomerBookNo,
    CustomerPageNo,
    CustomerOwe,
    CustomerLastProcess,
    CustomerRegisterDate,
    CustomerIsActive
) VALUES (
    @p1,
    @p2,
    @p3,
    @p4,
    @p5,
    @p6,
    @p7,
    @p8,
    @p9,
    @p10
)";

            using (OleDbConnection connection = new OleDbConnection(cs))
            {
                connection.Open();

                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];

                    string[] adSoyad = AdSoyadAyir(row[0].ToString());
                    string ad = adSoyad[0];
                    string soyad = adSoyad[1];
                    string bp = row[1].ToString();
                    string[] kitapSayfa = KitapSayfaAyir(bp);
                    string kitap = kitapSayfa[0];
                    string sayfa = kitapSayfa[1];
                    string telefonEv = row[2].ToString();
                    string telefonCep = row[3].ToString();
                    string hesapKapandi = row[4].ToString();

                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = query;

                        command.Parameters.AddWithValue("@p1", ad);
                        command.Parameters.AddWithValue("@p2", soyad);
                        command.Parameters.AddWithValue("@p3", bp + " - " + telefonCep + " - " + hesapKapandi);
                        command.Parameters.AddWithValue("@p4", telefonEv);
                        command.Parameters.AddWithValue("@p5", kitap);
                        command.Parameters.AddWithValue("@p6", sayfa);
                        command.Parameters.AddWithValue("@p7", hesapKapandi);
                        command.Parameters.AddWithValue("@p8", "Excelden aktarıldı.");
                        command.Parameters.AddWithValue("@p9", "02.02.2019 00:00");
                        command.Parameters.AddWithValue("@p10", "true");

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        private string[] AdSoyadAyir(string adSoyad)
        {
            string[] split = adSoyad.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (split.Length < 2) return new string[] { adSoyad, "" };

            string ad = string.Join(" ", split.Take(split.Length - 1));

            string soyad = split[split.Length - 1];

            return new string[] { ad, soyad };
        }

        private string[] KitapSayfaAyir(string kitapSayfa)
        {
            if (string.IsNullOrWhiteSpace(kitapSayfa)) return new string[] { string.Empty, string.Empty };
            if (kitapSayfa == "18123") return new string[] { "18", "123" };

            return kitapSayfa.Replace(' ', '-').Split('-');
        }

        #endregion
    }
}
