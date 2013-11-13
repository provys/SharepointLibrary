using System;
using System.Data;
using System.Management.Instrumentation;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

namespace SharepointLibrary
{
    public struct ReturnCode
    {
        public bool Result;
        public int ErrorCode;
        public string ErrorMessage;
        public Exception ErrorException;
    }

    public class SharepointAccess
    {
        private const string ОшибкаСписки = "Не удалось загрузить коллекцию списков с портала",
                             ОшибкаТекущийПользователь = "Не удалось загрузить информацию о текущем пользователе портала",
                             ОшибкаНетДанных = "Не удалось получить данные с портала";


        private readonly string _ссылкаПортала;
        private SP.ClientContext _портал;

        /// <summary>
        /// Инициализирует объект портала SharePoint и проверяет его доступность
        /// </summary>
        /// <param name="ссылкаПортала">Ссылка на список asmx портала SharePoint</param>
        public SharepointAccess(string ссылкаПортала)
        {
            _ссылкаПортала = ссылкаПортала;
            var проверка = ПроверитьДоступность();
            if (!проверка.Result) throw new EntryPointNotFoundException("#" + проверка.ErrorCode + ": " + проверка.ErrorMessage, проверка.ErrorException);
        }

        /// <summary>
        /// Проверяет доступность SharePoint сайта и его функционала
        /// </summary>
        /// <returns>Полная информация - была ли ошибка, какова она</returns>
        public ReturnCode ПроверитьДоступность()
        {
            try
            {
                _портал = new SP.ClientContext(_ссылкаПортала);
                var списки = ПолучитьСписки();
                if (списки == null || списки.Count == 0) throw new Exception(ОшибкаСписки);
            }
            catch (Exception ex) { return new ReturnCode() { Result = false, ErrorCode = -1, ErrorMessage = ex.Message, ErrorException = ex }; }
            return new ReturnCode() { Result = true, ErrorCode = 0, ErrorMessage = "", ErrorException = new Exception() };
        }

        /// <summary>
        /// Получает коллекцию списков с портала SharePoint. Если возник
        /// </summary>
        /// <returns>Коллекция списков портала. Требуется проверка на Exception и null</returns>
        public SP.ListCollection ПолучитьСписки()
        {
            try
            {
                if (_портал == null || _портал.Web == null) return null;
                SP.ListCollection списки = _портал.Web.Lists;
                _портал.Load(списки);
                _портал.ExecuteQuery();
                return списки;
            }
            catch (Exception ex) { throw new InstanceNotFoundException(ОшибкаСписки, ex); }
        }

        /// <summary>
        /// Получает всю информацию о текущем пользователе портала (по Windows-based аутентификации), от чьего имени осуществляется подключение
        /// </summary>
        /// <returns>ID пользователя, Email, login и т.д. текущего пользователь портала. Требуется проверка на Exception и null</returns>
        public SP.User ТекущийПользователь()
        {
            try
            {
                if (_портал == null || _портал.Web == null) return null;
                SP.User пользователь = _портал.Web.CurrentUser;
                _портал.Load(пользователь);
                _портал.ExecuteQuery();
                return пользователь;
            }
            catch (Exception ex) { throw new InstanceNotFoundException(ОшибкаТекущийПользователь, ex); }
        }

        /// <summary>
        /// Получает данные переданного списка, используя фильтр по колонкам строк
        /// </summary>
        /// <param name="список">Выбранный список, данные которого нужно получить.</param>
        /// <param name="условияВыборки">Условия фильтрации данных списка. Должны быть заключены в теги: Eq - равно<br/>Neq - не равно<br/>Gt - больше, чем...<br/>Geq - больше или равно<br/>Lt - меньше, чем...<br/>Leq - меньше или равно
        /// <br/>IsNull<br/>IsNotNull<br/>BeginsWith - начинается с... (для текста)<br/>Contains - содержит (для текста)<br/>Includes - включает</param>
        /// <param name="максимумСтрок">Максимальное количество строк, которое необходимо возвратить.</param>
        /// <returns>Данные списка</returns>
        public ListItemCollection ПолучитьДанныеСписка(SP.List список, string условияВыборки = "", int максимумСтрок = 3000)
        {
            try
            {
                if (_портал == null || _портал.Web == null) return null;

                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = string.Format("<View><Query><Where>{0}</Where></Query><RowLimit>{1}</RowLimit></View>", условияВыборки, максимумСтрок) /*<Geq><FieldRef Name='ID'/><Value Type='Number'>10</Value></Geq>*/ 
                };
                ListItemCollection данные = список.GetItems(camlQuery);

                _портал.Load(данные);
                _портал.ExecuteQuery();
                return данные;
            }
            catch (Exception ex) { throw new DataException(ОшибкаНетДанных, ex); }
        }



    }
}
