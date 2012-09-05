using System.Collections.Generic;
using System.IO;
using System.Linq;
using LinqToExcel;
using Raven.Abstractions.Data;
using Raven.Client;
using Raven.Client.Document;
using Raven.Client.Extensions;
using Raven.Json.Linq;

namespace ExcelToRavenImporter
{
    internal class ExcelImporter
    {
        private readonly string _excelFileName;
        private readonly string _ravenUrl;
        private readonly string _ravenDatabaseName;
        private DocumentStore _store;
        private IDocumentSession _session;
        private ExcelQueryFactory _excelQueryFactory;

        public ExcelImporter(string excelFileName, string ravenUrl)
        {
            _excelFileName = excelFileName;
            _ravenUrl = ravenUrl;
            _ravenDatabaseName = Path.GetFileNameWithoutExtension(excelFileName);
        }

        public void Execute()
        {
            _excelQueryFactory = new ExcelQueryFactory(_excelFileName);

            using (_store = new DocumentStore { Url = _ravenUrl })
            {
                _store.Initialize();
                _store.DatabaseCommands.EnsureDatabaseExists(_ravenDatabaseName);

                using (_session = _store.OpenSession(_ravenDatabaseName))
                {
                    foreach (string worksheetName in _excelQueryFactory.GetWorksheetNames())
                        ImportWorksheet(worksheetName);

                    ResetDocumentKeyGenerationToDefault();
                }
            }
        }

        private void ImportWorksheet(string worksheetName)
        {
            SetDocumentKeyPrefix(worksheetName);
            IList<string> columns = _excelQueryFactory.GetColumnNames(worksheetName).ToList();

            foreach (Row row in _excelQueryFactory.Worksheet(worksheetName))
            {
                RavenJObject doc = new RavenJObject();
                foreach (string column in columns)
                {
                    if (column.ToLower() == "id")
                        doc["Id"] = string.Format("{0}/{1}", worksheetName, row[column]);
                    else
                        doc.Add(column, new RavenJValue(row[column]));
                }
                _session.Store(doc);
                _session.Advanced.GetMetadataFor(doc)[Constants.RavenEntityName] = worksheetName;
            }
            _session.SaveChanges();
        }

        private void ResetDocumentKeyGenerationToDefault()
        {
            _store.Conventions.DocumentKeyGenerator = entity => _store.Conventions.GetTypeTagName(entity.GetType()) + "/";
        }

        private void SetDocumentKeyPrefix(string worksheetName)
        {
            _store.Conventions.DocumentKeyGenerator = entity => worksheetName + "/";
        }
    }
}