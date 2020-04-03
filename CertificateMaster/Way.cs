using System;

namespace CertificateMaster
{
    class Way
    {
        /// <summary>
        /// Путь к основной системной папке
        /// </summary>
        public static readonly string MainWay = string.Format(
            "{0}\\Мастер удостоверений\\",
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));

        /// <summary>
        /// Путь к настройкам основных ФИО
        /// </summary>
        public static string SettingsFIO = MainWay + "Данные\\SettingsFIO.xml";

        /// <summary>
        /// Путь к последней введённой информации
        /// </summary>
        public static string LastData = MainWay + "Данные\\LastData.xml";

        /// <summary>
        /// Путь к шаблону удостоверения
        /// </summary>
        public static string Certificate1 = MainWay + "Шаблоны\\Certificate1.doc";

        /// <summary>
        /// Путь к шаблону удостоверения
        /// </summary>
        public static string Certificate2 = MainWay + "Шаблоны\\Certificate2.doc";

    }

}


