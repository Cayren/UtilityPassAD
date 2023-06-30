namespace PasswordExpiration
{
    class Program
    {
        static void Main(string[] args)
        {
            List<User> usuarios = GetADUsers();
            DateTime fechaExpiracion;
            int dias = 0;
            foreach (User usuario in usuarios)
            {
                fechaExpiracion = usuario.PasswordExpire;
                if (fechaExpiracion != null && fechaExpiracion.ToString() != "01/01/0001 0:00:00")
                {
                    //restar fecha a hoy y si la diferencia da 7 o menos hay que mandar un correo
                    dias = (fechaExpiracion - DateTime.Today).Days;
                    if (dias <= 7)
                    {
                        Enviar(usuario.Email, dias);
                    }
                }
            }
        }

        private static void Enviar(string cuenta, int dias)
        {

          //configurar servicio de correo
            MailMessage mensaje = new MailMessage();
            SmtpClient smtp = new SmtpClient();

            MailAddress correo = new MailAddress(cuenta);
            mensaje.To.Add(correo);

            mensaje.From = new MailAddress("correo@gmail.com");
            mensaje.Subject = "CADUCIDAD CONTRASEÑA";
            mensaje.Body = "Tu contraseña caduca en " + dias;
            mensaje.IsBodyHtml = true;
            mensaje.Priority = MailPriority.High;

            smtp.Port = 465;
            smtp.Host = "smtp-relay.gmail.com";
            smtp.EnableSsl = false;
            smtp.UseDefaultCredentials = false;
            NetworkCredential credenciales = new NetworkCredential("correo@gmail.com", "pass");
            smtp.Credentials = credenciales;
            smtp.Send(mensaje);

            mensaje.Dispose();

        }

        public static List<User> GetADUsers()
        {
            List<User> rst = new List<User>();

            string DomainPath = "LDAP://DC=xxxxx,DC=xxxx";
            DirectoryEntry adSearchRoot = new DirectoryEntry(DomainPath);
            DirectorySearcher adSearcher = new DirectorySearcher(adSearchRoot);

            adSearcher.Filter = "(&(objectClass=user)(objectCategory=person))";
            adSearcher.PropertiesToLoad.Add("samaccountname");
            adSearcher.PropertiesToLoad.Add("title");
            adSearcher.PropertiesToLoad.Add("mail");
            adSearcher.PropertiesToLoad.Add("usergroup");
            adSearcher.PropertiesToLoad.Add("company");
            adSearcher.PropertiesToLoad.Add("department");
            adSearcher.PropertiesToLoad.Add("telephoneNumber");
            adSearcher.PropertiesToLoad.Add("pwdLastSet");
            adSearcher.PropertiesToLoad.Add("msDS-UserPasswordExpiryTimeComputed");
            adSearcher.PropertiesToLoad.Add("displayname");
            SearchResult result;
            SearchResultCollection iResult = adSearcher.FindAll();

            User item;
            if (iResult != null)
            {
                for (int counter = 0; counter < iResult.Count; counter++)
                {
                    result = iResult[counter];
                    if (result.Properties.Contains("samaccountname"))
                    {
                        item = new User();

                        item.UserName = (String)result.Properties["samaccountname"][0];

                        if (result.Properties.Contains("displayname"))
                        {
                            item.DisplayName = (String)result.Properties["displayname"][0];
                        }

                        if (result.Properties.Contains("mail"))
                        {
                            item.Email = (String)result.Properties["mail"][0];
                        }

                        if (result.Properties.Contains("company"))
                        {
                            item.Company = (String)result.Properties["company"][0];
                        }

                        if (result.Properties.Contains("title"))
                        {
                            item.JobTitle = (String)result.Properties["title"][0];
                        }

                        if (result.Properties.Contains("department"))
                        {
                            item.Deparment = (String)result.Properties["department"][0];
                        }

                        if (result.Properties.Contains("telephoneNumber"))
                        {
                            item.Phone = (String)result.Properties["telephoneNumber"][0];
                        }

                        if (result.Properties.Contains("pwdLastSet"))
                        {
                            long valorprueba = (long)result.Properties["pwdLastSet"][0];
                            item.PasswordLastSet = DateTime.FromFileTimeUtc(valorprueba).ToString();
                        }

                        if (result.Properties.Contains("msDS-UserPasswordExpiryTimeComputed"))
                        {
                            long valor = (long)result.Properties["msDS-UserPasswordExpiryTimeComputed"][0];
                            if (valor != 9223372036854775807)
                            {
                                item.PasswordExpire = DateTime.FromFileTimeUtc(valor);
                            }
                        }
                        rst.Add(item);
                    }
                }
            }

            adSearcher.Dispose();
            adSearchRoot.Dispose();

            return rst;
        }

        public class User
        {
            public string UserName { get; set; }

            public string DisplayName { get; set; }

            public string Company { get; set; }

            public string Deparment { get; set; }

            public string JobTitle { get; set; }

            public string Email { get; set; }

            public string Phone { get; set; }

            public string PasswordLastSet { get; set; }

            public DateTime PasswordExpire { get; set; }
        }
    }
}
