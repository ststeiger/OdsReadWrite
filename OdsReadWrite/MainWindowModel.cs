using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using OdsReadWrite.Properties;
using System.Diagnostics;

namespace OdsReadWrite
{
    internal sealed class MainWindowModel : INotifyPropertyChanged, IDataErrorInfo
    {
        private const string
            DataPropertyName = "Data",
            SheetsPropertyName = "Sheets",
            RowsPropertyName = "Rows",
            ColumnsPropertyName = "Columns",
            InputPathPropertyName = "InputPath",
            OutputPathPropertyName = "OutputPath",
            OpenCreatedFilePropertyName = "OpenCreatedFile",
            SelectedDataTableIndexPropertyName = "SelectedDataTableIndex";

        private readonly DelegateCommand createCommand;
        private readonly DelegateCommand readCommand;
        private readonly DelegateCommand writeCommand;

        private readonly DelegateCommand showOpenFileDialogCommand;
        private readonly DelegateCommand showSaveFileDialogCommand;

        private DataSet data;

        private string sheets = Settings.Default.Sheets.ToString();
        private string rows = Settings.Default.Rows.ToString();
        private string columns = Settings.Default.Columns.ToString();

        private string inputPath = Settings.Default.InputPath;
        private string outputPath = Settings.Default.OutputPath;

        private bool openCreatedFile = Settings.Default.OpenCreatedFile;

        private int selectedDataTableIndex;

        public event PropertyChangedEventHandler PropertyChanged;

        private static int? ToPositiveIntegerOrNull(string value)
        {
            int result;
            if (int.TryParse(value, out result) && result > 0)
                return result;
            else
                return null;
        }

        private static bool DirectoryExists(string path)
        {
            return string.IsNullOrWhiteSpace(path) ? false : Directory.Exists(Path.GetDirectoryName(Path.GetFullPath(path)));
        }

        public MainWindowModel()
        {
            this.createCommand = new DelegateCommand(this.Create, this.CanCreate);
            this.readCommand = new DelegateCommand(this.Read, this.CanRead);
            this.writeCommand = new DelegateCommand(this.Write, this.CanWrite);

            this.showOpenFileDialogCommand = new DelegateCommand(this.ShowOpenFileDialog);
            this.showSaveFileDialogCommand = new DelegateCommand(this.ShowSaveFileDialog);

            this.Create();

            App.Current.Exit += this.OnExit;
        }

        public DelegateCommand CreateCommand
        {
            get { return this.createCommand; }
        }

        public DelegateCommand ReadCommand
        {
            get { return this.readCommand; }
        }

        public DelegateCommand WriteCommand
        {
            get { return this.writeCommand; }
        }

        public DelegateCommand ShowOpenFileDialogCommand
        {
            get { return this.showOpenFileDialogCommand; }
        }

        public DelegateCommand ShowSaveFileDialogCommand
        {
            get { return this.showSaveFileDialogCommand; }
        }

        public DataSet Data
        {
            get { return this.data; }
            private set { this.SetValue(DataPropertyName, ref this.data, value); }
        }

        public string Sheets
        {
            get { return this.sheets; }
            set { this.SetValue(SheetsPropertyName, ref this.sheets, value); }
        }

        public string Rows
        {
            get { return this.rows; }
            set { this.SetValue(RowsPropertyName, ref this.rows, value); }
        }

        public string Columns
        {
            get { return this.columns; }
            set { this.SetValue(ColumnsPropertyName, ref this.columns, value); }
        }

        private int? SheetsCount
        {
            get { return ToPositiveIntegerOrNull(this.Sheets); }
        }

        private int? RowsCount
        {
            get { return ToPositiveIntegerOrNull(this.Rows); }
        }

        private int? ColumnsCount
        {
            get { return ToPositiveIntegerOrNull(this.Columns); }
        }

        public string InputPath
        {
            get { return this.inputPath; }
            set { this.SetValue(InputPathPropertyName, ref this.inputPath, value); }
        }

        public string OutputPath
        {
            get { return this.outputPath; }
            set { this.SetValue(OutputPathPropertyName, ref this.outputPath, value); }
        }

        public bool OpenCreatedFile
        {
            get { return this.openCreatedFile; }
            set { this.SetValue(OpenCreatedFilePropertyName, ref this.openCreatedFile, value); }
        }

        public int SelectedDataTableIndex
        {
            get { return this.selectedDataTableIndex; }
            set { this.SetValue(SelectedDataTableIndexPropertyName, ref this.selectedDataTableIndex, value); }
        }

        public string Error
        {
            get { return string.Empty; }
        }

        public string this[string propertyName]
        {
            get
            {
                switch (propertyName)
                {
                    case SheetsPropertyName:
                        this.CreateCommand.InvalidateCanExecuteChanged();
                        return this.SheetsCount.HasValue ? string.Empty : "Value must be a positive integer.";
                    case RowsPropertyName:
                        this.CreateCommand.InvalidateCanExecuteChanged();
                        return this.RowsCount.HasValue ? string.Empty : "Value must be a positive integer.";
                    case ColumnsPropertyName:
                        this.CreateCommand.InvalidateCanExecuteChanged();
                        return this.ColumnsCount.HasValue ? string.Empty : "Value must be a positive integer.";
                    case InputPathPropertyName:
                        this.ReadCommand.InvalidateCanExecuteChanged();
                        return File.Exists(this.InputPath) ? string.Empty : "File doesn't exist.";
                    case OutputPathPropertyName:
                        this.WriteCommand.InvalidateCanExecuteChanged();
                        return DirectoryExists(this.OutputPath) ? string.Empty : "Directory doesn't exist.";
                    default:
                        return string.Empty;
                }
            }
        }

        private void SetValue<T>(string propertyName, ref T field, T value)
        {
            if (!EqualityComparer<T>.Default.Equals(field, value))
            {
                field = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs(propertyName));
            }
        }

        private void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            var handler = this.PropertyChanged;
            if (handler != null)
                handler(this, e);
        }

        private void OnExit(object sender, ExitEventArgs e)
        {
            var count = this.SheetsCount;
            if (count.HasValue)
                Settings.Default.Sheets = count.Value;

            count = this.RowsCount;
            if (count.HasValue)
                Settings.Default.Rows = count.Value;

            count = this.ColumnsCount;
            if (count.HasValue)
                Settings.Default.Columns = count.Value;

            if (File.Exists(this.InputPath))
                Settings.Default.InputPath = this.InputPath;
            if (DirectoryExists(this.OutputPath))
                Settings.Default.OutputPath = this.OutputPath;

            Settings.Default.OpenCreatedFile = this.OpenCreatedFile;

            Settings.Default.Save();
        }

        private bool CanCreate()
        {
            return this.SheetsCount.HasValue && this.RowsCount.HasValue && this.ColumnsCount.HasValue;
        }

        private void Create()
        {
            if (this.CanCreate())
            {
                var data = new DataSet();

                for (int i = 0, sheetsCount = this.SheetsCount.Value; i < sheetsCount; ++i)
                {
                    var table = data.Tables.Add(string.Format(CultureInfo.InvariantCulture, "Sheet {0}", i + 1));
                    int columnsCount = this.ColumnsCount.Value;
                    for (int j = 0; j < columnsCount; ++j)
                        table.Columns.Add(string.Format(CultureInfo.InvariantCulture, "Column {0}", j + 1), typeof(string));

                    for (int j = 0, rowsCount = this.RowsCount.Value; j < rowsCount; ++j)
                        table.Rows.Add(new string[columnsCount]);
                }

                this.Data = data;
                this.SelectedDataTableIndex = 0;

                this.WriteCommand.InvalidateCanExecuteChanged();
            }
        }

        private bool CanRead()
        {
            return File.Exists(this.InputPath);
        }

        private void Read()
        {
            if (this.CanCreate())
            {
                this.Data = new OdsReaderWriter().ReadOdsFile(this.InputPath);
                this.SelectedDataTableIndex = 0;
            }
        }

        private bool CanWrite()
        {
            return this.Data != null && DirectoryExists(this.OutputPath);
        }

        private void Write()
        {
            if (this.CanWrite())
            {
                new OdsReaderWriter().WriteOdsFile(this.Data, this.OutputPath);

                if (this.OpenCreatedFile)
                    Process.Start(this.OutputPath);
            }
        }

        private void ShowOpenFileDialog()
        {
            var openFileDialog = new OpenFileDialog()
            {
                FileName = this.InputPath,
                DefaultExt = "*.ods",
                Filter = "Open Document Spreadsheet (.ods)|*.ods"
            };

            if (openFileDialog.ShowDialog() == true)
                this.InputPath = openFileDialog.FileName;
        }

        private void ShowSaveFileDialog()
        {
            var saveFileDialog = new SaveFileDialog()
            {
                FileName = this.OutputPath,
                DefaultExt = "*.ods",
                Filter = "Open Document Spreadsheet (.ods)|*.ods"
            };

            if (saveFileDialog.ShowDialog() == true)
                this.OutputPath = saveFileDialog.FileName;
        }
    }
}
