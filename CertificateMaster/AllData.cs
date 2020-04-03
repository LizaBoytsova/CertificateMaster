using System;
using System.Drawing.Printing;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;

namespace CertificateMaster
{
    [Serializable]
    public class AllData
    {
        #region Общее

        /// <summary>
        /// Количество сертификатов на листе
        /// </summary>
        public int CountCert;

        /// <summary>
        /// Печать без предварительного просмотра
        /// </summary>
        public bool justPrint;

        /// <summary>
        /// Принтер для печати
        /// </summary>
        public string PrintName;

        #endregion Общее

        #region Сертификат 1

        /// <summary>
        /// Номер удостоверения
        /// </summary>
        public int CertNumTab1;

        /// <summary>
        /// ФИО
        /// </summary>
        public string FIOTab1;

        /// <summary>
        /// Должность1
        /// </summary>
        public string Post1Tab1;

        /// <summary>
        /// Должность2
        /// </summary>
        public string Post2Tab1;

        /// <summary>
        /// Место работы 1
        /// </summary>
        public string Workplace1Tab1;

        /// <summary>
        /// Место работы 2
        /// </summary>
        public string Workplace2Tab1;

        /// <summary>
        /// Дата начала обучения
        /// </summary>
        public DateTime dateStartLearningTab1;

        /// <summary>
        /// Дата окончания обучения
        /// </summary>
        public DateTime dateEndLearningTab1;

        /// <summary>
        /// Обучения для кого
        /// </summary>
        public string LearningForWhoTab1;

        /// <summary>
        /// Номер протокола
        /// </summary>
        public int ProtocolNumberTab1;

        /// <summary>
        /// Дата протокола
        /// </summary>
        public DateTime dateProtocolTab1;

        /// <summary>
        /// Дата окончания действия удостоверения
        /// </summary>
        public DateTime dateEndCertificateTab1;

        #endregion Сертификат 1


        #region Сертификат 2

        /// <summary>
        /// Номер удостоверения
        /// </summary>
        public int CertNumTab2;

        /// <summary>
        /// ФИО
        /// </summary>
        public string FIOTab2;

        /// <summary>
        /// Должность1
        /// </summary>
        public string Post1Tab2;

        /// <summary>
        /// Должность2
        /// </summary>
        public string Post2Tab2;

        /// <summary>
        /// Место работы 1
        /// </summary>
        public string Workplace1Tab2;

        /// <summary>
        /// Место работы 2
        /// </summary>
        public string Workplace2Tab2;

        /// <summary>
        /// Дата начала обучения
        /// </summary>
        public DateTime dateStartLearningTab2;

        /// <summary>
        /// Дата окончания обучения
        /// </summary>
        public DateTime dateEndLearningTab2;

        /// <summary>
        /// Обучения для кого
        /// </summary>
        public string LearningForWhoTab2;

        /// <summary>
        /// Номер протокола
        /// </summary>
        public int ProtocolNumberTab2;

        /// <summary>
        /// Дата протокола
        /// </summary>
        public DateTime dateProtocolTab2;

        /// <summary>
        /// Дата окончания действия удостоверения
        /// </summary>
        public DateTime dateEndCertificateTab2;

        #endregion Сертификат 3

    }

}
