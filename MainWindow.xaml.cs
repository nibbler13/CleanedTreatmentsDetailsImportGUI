using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CleanedTreatmentsDetailsImport;
using System.Diagnostics;
using System.Data;
using System.IO;
using System.Threading;
using System.Windows.Threading;
using System.Windows.Documents;

namespace DataImportToDB {
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window, INotifyPropertyChanged {
		public event PropertyChangedEventHandler PropertyChanged;
		private void NotifyPropertyChanged([CallerMemberName] String propertyName = "") {
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}

		private BackgroundWorker bwReadFile;
		private BackgroundWorker bwLoadToDb;
		private readonly DispatcherTimer timerWorkIndicator;
		private string workIndicator = "###";

		private readonly Dictionary<string, string> insuranceCompaniesJID = new Dictionary<string, string> {
			{ "АбсолютСтрахование", "991515382,991519409,991519865,991523030" },
			{ "Альфастрахование", "100005,991520911,991514852" },
			{ "Альфастрахование Санкт-Петербург", "990424275" },
			{ "Альянс", "991511535,991520499,991519440,991521374,991511568" },
			{ "БестДоктор", "991520964, 991526106" },
			{ "ВСК", "991516556,991520387,991523215,991519361,991525970" },
			{ "ВТБ", "991515797,991520427" },
			{ "Другие СК", "" },
			{ "Ингосстрах взр", "991522348,991522924,991525955,991522442" },
			{ "Ингосстрах дет", "991522386" },
			{ "ЛибертиСтрахование", "991517912" },
			{ "Метлайф", "991517927,991523451,991519436" },
			{ "Росгосстрах", "991511705,1990097479" },
			{ "Ренессанс", "991523042,991523280,991523170" },
			{ "СК РЕСО", "991518370,991521272,991523038,991526075,991519595" },
			{ "СМП страхование", "991516698,991521960" },
			{ "СОГАЗ", "991524638" },
			{ "Согласие", "991520913,991518470,991519761,991523028" },
			{ "Энергогарант", "991523453,991517214,991520348" },
			{ "Ингосстрах Сочи", "991512906" },
			{ "Ингосстрах Краснодар", "991357338" },
			{ "Ингосстрах Уфа", "991370062" },
			{ "Ингосстрах Санкт-Петербург", "990389345" },
			{ "Ингосстрах Казань", "991379370" },
			{ "БестДоктор Санкт-Петербург", "991523486" },
			{ "БестДоктор Уфа", "991523489" },
			{ "СОГАЗ Уфа", "991524671,991524697" }
		};

		public List<string> InsuranceCompanies { get; set; } = new List<string>();

		private string selectedInsuranceCompany;
		public string SelectedInsuranceCompany { 
			get { return selectedInsuranceCompany; }
			set {
				if (value != selectedInsuranceCompany) {
					selectedInsuranceCompany = value;
					NotifyPropertyChanged();
				}
			}
		}


		private DateTime dateBegin;
		public DateTime DateBegin { 
			get { return dateBegin; }
			set {
				if (value != dateBegin) {
					dateBegin = value;
					NotifyPropertyChanged();
				}
			}
		}

		private DateTime dateEnd;
		public DateTime DateEnd { 
			get { return dateEnd; }
			set {
				if (value != dateEnd) {
					dateEnd = value;
					NotifyPropertyChanged();
				}
			}
		}

		private string selectedFile;
		public string SelectedFile { 
			get { return string.IsNullOrEmpty(selectedFile) ? string.Empty : Path.GetFileName(selectedFile); }
			set {
				if (value != selectedFile) {
					selectedFile = value;
					NotifyPropertyChanged();
				}
			}
		}

		private string progressText;
		public string ProgressText {
			get { return progressText; }
			set {
				if (value != progressText) {
					progressText = value;
					NotifyPropertyChanged();
				}
			}
		}

		public ObservableCollection<string> SheetNames { get; set; } = new ObservableCollection<string>();

		private string selectedSheetName;
		public string SelectedSheetName { 
			get { return selectedSheetName; }
			set {
				if (value != selectedSheetName) {
					selectedSheetName = value;
					NotifyPropertyChanged();
				}
			}
		}


		private bool isButtonLoadToDbEnabled;
		public bool IsButtonLoadToDbEnabled {
			get { return isButtonLoadToDbEnabled; }
			set {
				if (value != isButtonLoadToDbEnabled) {
					isButtonLoadToDbEnabled = value;
					NotifyPropertyChanged();
				}
			}
		}

		private string fileInfo = "unknown";

		private bool isButtonReadFileEnabled = true;
		public bool IsButtonReadFileEnabled {
			get { return isButtonReadFileEnabled; }
			set {
				if (value != isButtonReadFileEnabled) {
					isButtonReadFileEnabled = value;
					NotifyPropertyChanged();
				}
			}
		}


		private bool isCheckedloadTypeTreatmentsDetails;
		public bool IsCheckedLoadTypeTreatmentsDetails {
			get { return isCheckedloadTypeTreatmentsDetails; }
			set {
				if (value != isCheckedloadTypeTreatmentsDetails) {
					isCheckedloadTypeTreatmentsDetails = value;
					NotifyPropertyChanged();
					SetUpVisibilityForLoadType();
				}
			}
		}

		private bool isCheckedloadTypeProfitAndLoss;
		public bool IsCheckedLoadTypeProfitAndLoss {
			get { return isCheckedloadTypeProfitAndLoss; }
			set {
				if (value != isCheckedloadTypeProfitAndLoss) {
					isCheckedloadTypeProfitAndLoss = value;
					NotifyPropertyChanged();
					SetUpVisibilityForLoadType();
				}
			}
		}

		private void SetUpVisibilityForLoadType() {
			VisibilityInsuranceCompanyComboBox = IsCheckedLoadTypeTreatmentsDetails ? Visibility.Visible : Visibility.Collapsed;
			VisibilityPeriodGroupBox = IsCheckedLoadTypeTreatmentsDetails ? Visibility.Visible : Visibility.Collapsed;
		}


		private Visibility visibilityInsuranceCompanyComboBox;
		public Visibility VisibilityInsuranceCompanyComboBox {
			get { return visibilityInsuranceCompanyComboBox; }
			set {
				if (value != visibilityInsuranceCompanyComboBox) {
					visibilityInsuranceCompanyComboBox = value;
					NotifyPropertyChanged();
				}
			}
		}


		private Visibility visibilityPeriodGroupBox;
		public Visibility VisibilityPeriodGroupBox {
			get { return visibilityPeriodGroupBox; }
			set {
				if (value != visibilityPeriodGroupBox) {
					visibilityPeriodGroupBox = value;
					NotifyPropertyChanged();
				}
			}
		}




		public MainWindow() {
			InitializeComponent();

			InsuranceCompanies.AddRange(insuranceCompaniesJID.Keys);
			DateBegin = DateTime.Now.Date.AddDays(DateTime.Now.Day * -1);
			DateEnd = DateBegin;
			DateBegin = DateBegin.AddDays((DateBegin.Day - 1) * -1);
			IsCheckedLoadTypeTreatmentsDetails = true;

			DataContext = this;

			//Loaded += (s, e) => {
			//	if (Debugger.IsAttached) {
			//		SelectedFile = @"C:\Users\nn-admin\Desktop\PL шаблон.xlsx";
			//		IsCheckedLoadTypeProfitAndLoss = true;
			//		ReadSheetNames();
			//	}
			//};

			Closing += (s, e) => {
				if ((bwReadFile != null && bwReadFile.IsBusy) ||
					(bwLoadToDb != null && bwLoadToDb.IsBusy)) {
					if (MessageBox.Show(
						this,
						"Текущие операции еще не завершены" + Environment.NewLine +
						"Данные могут быть потеряны" + Environment.NewLine + 
						"Вы уверены, что хотите завершить работу?",
						string.Empty,
						MessageBoxButton.YesNo,
						MessageBoxImage.Question) == MessageBoxResult.No) {
						e.Cancel = true;
					}
				}
			};

			timerWorkIndicator = new DispatcherTimer();
			timerWorkIndicator.Interval = TimeSpan.FromMilliseconds(300);
			timerWorkIndicator.Tick += (s, e) => {
				UpdateProgress(string.Empty, true);
			};
		}







		private void ButtonSelectFile_Click(object sender, RoutedEventArgs e) {
			SheetNames.Clear();
			SelectedFile = string.Empty;

			OpenFileDialog openFileDialog = new OpenFileDialog {
				Filter = "Книга Excel (*.xls*)|*.xls*",
				CheckFileExists = true,
				CheckPathExists = true,
				Multiselect = false,
				RestoreDirectory = true
			};

			if (openFileDialog.ShowDialog() == true) {
				SelectedFile = openFileDialog.FileName;
				ReadSheetNames();
			}
		}

		private void ReadSheetNames() {
			SheetNames.Clear();
			SelectedSheetName = string.Empty;

			try {
				ExcelReader.ReadSheetNames(selectedFile).ForEach(SheetNames.Add);

				if (SheetNames.Count > 0) {
					if (SheetNames.Contains("Лист1"))
						SelectedSheetName = "Лист1";
					else if (SheetNames.Contains("Данные"))
						SelectedSheetName = "Данные";
					else
						SelectedSheetName = SheetNames[0];
				}
			} catch (Exception exc) {
				MessageBox.Show(
					this,
					exc.Message + Environment.NewLine + exc.StackTrace,
					"Ошибка считывания Excel файла",
					MessageBoxButton.OK,
					MessageBoxImage.Error);
			}
		}

		private void ButtonReadSelectedFile_Click(object sender, RoutedEventArgs e) {
			string error = string.Empty;

			if (string.IsNullOrEmpty(SelectedFile)) {
				error = "Не выбран файл с данными";
			} else if (string.IsNullOrEmpty(SelectedInsuranceCompany)) {
				if (IsCheckedLoadTypeTreatmentsDetails)
					error = "Не выбрана страховая компания";
			} else if (string.IsNullOrEmpty(SelectedSheetName)) {
				error = "Не выбран лист в файле";
			} else if (DateBegin > DateEnd) {
				if (IsCheckedLoadTypeTreatmentsDetails)
					error = "Дата начала периода больше даты окончания";
			} else if (DateBegin > DateTime.Now.Date) {
				if (IsCheckedLoadTypeTreatmentsDetails)
					error = "Дата начала периода больше текущего дня";
			}

			if (!string.IsNullOrEmpty(error)) {
				MessageBox.Show(
					this,
					error,
					string.Empty,
					MessageBoxButton.OK,
					MessageBoxImage.Warning);
				return;
			}

			ProgressText = string.Empty;

			Cursor = Cursors.Wait;
			IsButtonReadFileEnabled = false;
			StartWorkIndicator();

			bwReadFile = new BackgroundWorker();
			bwReadFile.DoWork += BwReadFile_DoWork;
			bwReadFile.ProgressChanged += BwReadFile_ProgressChanged;
			bwReadFile.WorkerReportsProgress = true;
			bwReadFile.RunWorkerCompleted += BwReadFile_RunWorkerCompleted;
			bwReadFile.RunWorkerAsync();
		}

		private void BwReadFile_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
			Cursor = Cursors.Arrow;
			IsButtonReadFileEnabled = true;
			StopWorkIndicator();
			bool result = e.Result != null && ((bool)e.Result == true);
			IsButtonLoadToDbEnabled = result;
			UpdateProgress(new string('=', 40));

			if (e.Error != null) {
				string messageError = "Ошибка выполнения: " + Environment.NewLine + 
					Environment.NewLine + e.Error.Message + Environment.NewLine + e.Error.StackTrace;
				Logging.ToLog(messageError);
				UpdateProgress(messageError);
				MessageBox.Show(
					this,
					messageError,
					string.Empty,
					MessageBoxButton.OK,
					MessageBoxImage.Error);
				return;
			}

			if (!result) {
				string messageError = "Завершено" + Environment.NewLine +
					"Имеются проблемы из-за которых загрузка файла невозможна";
				Logging.ToLog(messageError);
				UpdateProgress(messageError);
				MessageBox.Show(
					this,
					messageError,
					string.Empty,
					MessageBoxButton.OK,
					MessageBoxImage.Warning);
				return;
			}

			string message = "Завершено успешно!" + Environment.NewLine + Environment.NewLine +
				"Можно выполнить загрузку в БД";
			Logging.ToLog(message);
			UpdateProgress(message);
			MessageBox.Show(
				this,
				message,
				string.Empty,
				MessageBoxButton.OK,
				MessageBoxImage.Information);
		}

		private void BwReadFile_ProgressChanged(object sender, ProgressChangedEventArgs e) {
			UpdateProgress(e.UserState.ToString());
		}



		private void UpdateProgress(string text, bool isWorkIndicator = false) {
			if (string.IsNullOrEmpty(text) && !isWorkIndicator)
				return;

			Application.Current.Dispatcher.Invoke(new Action(() => {
				ProgressText = ProgressText.Replace(workIndicator, "");

				if (isWorkIndicator) {
					workIndicator += "#";

					if (workIndicator.Length > 80)
						workIndicator = "###";

					ProgressText += workIndicator;
					TextBoxProgress.ScrollToEnd();
				} else {
					ProgressText += DateTime.Now.ToString("HH:mm:ss") + ": " + text + Environment.NewLine;
					TextBoxProgress.ScrollToEnd();
					workIndicator = "###";
				}
			}));
		}



		private void BwReadFile_DoWork(object sender, DoWorkEventArgs e) {
			e.Result = false;
			BackgroundWorker bw = sender as BackgroundWorker;

			UpdateProgress("Считывание файла: " + SelectedFile + ", лист: " + SelectedSheetName);
			string fileResult;
			if (IsCheckedLoadTypeTreatmentsDetails)
				fileResult = Program.ReadTreatmentsDetailsFileContent(selectedFile, SelectedSheetName, out fileInfo, bw);
			else if (IsCheckedLoadTypeProfitAndLoss)
				fileResult = Program.ReadProfitAndLossLFileContent(selectedFile, SelectedSheetName, out fileInfo, bw);
			else
				fileResult = "Неизвестный тип импорта; ошибок: 1";

			UpdateProgress(fileResult);
			if (!fileResult.Contains("ошибок: 0")) {
				UpdateProgress("!!! Имеются ошибки при считывании файла, продолжение невозможно");

				if (!Debugger.IsAttached)
					return;
			}

			int rowsReadedCount = 0;
			if (IsCheckedLoadTypeTreatmentsDetails)
				rowsReadedCount = Program.FileContentTreatmentsDetails.Rows.Count;
			else if (IsCheckedLoadTypeProfitAndLoss)
				rowsReadedCount = Program.FileContentProfitAndLoss.Count;

			if (rowsReadedCount == 0) {
				UpdateProgress("!!! Не удалось считать данные для записи в БД (пустой файл / несоответствие формата))");
				return;
			}

			if (IsCheckedLoadTypeProfitAndLoss) {
				e.Result = true;
				return;
			}

			UpdateProgress("Создание подключения к БД");
			VerticaClient verticaClient = new VerticaClient(
				VerticaSettings.host,
				VerticaSettings.database,
				VerticaSettings.user,
				VerticaSettings.password,
				bw);

			UpdateProgress("Запрос данных по выбранной страховой за указанный период");
			string sqlQuery = VerticaSettings.sqlGetDataTreatmentsDetails;
			if (SelectedInsuranceCompany.Equals("Другие СК"))
				sqlQuery = VerticaSettings.sqlGetDataTreatmentsDetailsOtherIc;

			DataTable dataTableDb = verticaClient.GetDataTable(
				sqlQuery.Replace("@jids", insuranceCompaniesJID[SelectedInsuranceCompany]),
				new Dictionary<string, object> {
					{ "@dateBegin", DateBegin },
					{ "@dateEnd", DateEnd }
				});

			UpdateProgress("Получено строк: " + dataTableDb.Rows.Count);

			if (dataTableDb.Rows.Count == 0) {
				UpdateProgress("!!! Не удалось получить информацию из БД за указанный период (нет приемов)");
				return;
			}

			UpdateProgress(new string('-', 40));
			UpdateProgress("Сверка данных между файлом Excel и БД");
			bool isDataOk = Program.IsCompareReadedDataToDbOk(
				dataTableDb, out int rowsToDelete, out int rowsLoadedBefore, out int rowsWithNoOrdtid, bw);

			UpdateProgress("Расчет и применение средней скидки для изначальных данных");
			Program.ApplyAverageDiscount(dataTableDb, bw);

			UpdateProgress(new string('=', 40));

			if (isDataOk) {
				UpdateProgress("Сводная информация по загруженным данным: ");

				int serviceCountLoaded = 0;
				double serviceAmountLoaded = 0;
				int rowCountLoaded = 0;

				foreach (DataRow dataRow in Program.FileContentTreatmentsDetails.Rows) {
					string contract = dataRow["contract"].ToString();
					if (string.IsNullOrEmpty(contract))
						continue;

					int serviceCount = int.Parse(dataRow["service_count"].ToString());
					serviceCountLoaded += serviceCount;
					double serviceAmount = double.Parse(dataRow["amount_total_with_discount"].ToString());
					serviceAmountLoaded += serviceAmount;
					rowCountLoaded++;
				}

				int serviceCountDb = 0;
				double serviceAmountDb = 0;
				int rowCountDb = dataTableDb.Rows.Count;

				foreach (DataRow dataRow in dataTableDb.Rows) {
					int serviceCount = int.Parse(dataRow["schcount"].ToString());
					serviceCountDb += serviceCount;
					double serviceAmount = double.Parse(dataRow["schamount"].ToString());
					serviceAmountDb += serviceAmount;
				}

				UpdateProgress(
					"\t\t| Кол-во услуг\t| Общая стоимость\t| Кол-во строк" + Environment.NewLine +
					"Загруженный файл\t| " + serviceCountLoaded.ToString("N0") + "\t\t| " + serviceAmountLoaded.ToString("N2") + "\t\t| " + rowCountLoaded.ToString("N0") + Environment.NewLine +
					"Данные в БД\t\t| " + serviceCountDb.ToString("N0") + "\t\t| " + serviceAmountDb.ToString("N2") + "\t\t| " + rowCountDb.ToString("N0"));

				UpdateProgress(new string('-', 80));
				UpdateProgress("Кол-во строк, помеченных для удаления: " + rowsToDelete);
				UpdateProgress(new string('-', 80));
				UpdateProgress("Сверка выполнена успешно, ошибок нет");
			} else {
				UpdateProgress("Считано строк из файла: " + rowsReadedCount);

				if (rowsLoadedBefore > 0)
					UpdateProgress("!!! Строк, загруженных ранее: " + rowsLoadedBefore + 
						" - не допускается повторная загрузка данных, необходимо сначала удалить ранее загруженные данные");

				if (rowsWithNoOrdtid > 0) {
					UpdateProgress("!!! Строк, для которых не удалось найти сопоставления в БД: " + rowsWithNoOrdtid +
						" - не допускается загрузка данных, которые невозможно идентифицировать в БД");

					if (Debugger.IsAttached)
						isDataOk = true;
				}

				UpdateProgress("Для консультирования просьба обратиться в отдел поддержки бизнес-приложений");
			}

			e.Result = isDataOk;
		}

		private void ButtonDateSelect_Click(object sender, RoutedEventArgs e) {
			string param = (sender as Button).Tag.ToString();

			if (param.Equals("EquateEndDateToBeginDate")) {
				DateEnd = DateBegin;
			} else if (param.Equals("SetDatesToCurrentDay")) {
				DateBegin = DateTime.Now;
				DateEnd = DateBegin;
			} else if (param.Equals("SetDatesToCurrentWeek")) {
				DateEnd = DateTime.Now;
				int dayOfWeek = (int)DateEnd.DayOfWeek;
				if (dayOfWeek == 0)
					dayOfWeek = 7;
				DateBegin = DateEnd.AddDays(-1 * (dayOfWeek - 1));
			} else if (param.Equals("SetDatesToCurrentMonth")) {
				DateBegin = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
				DateEnd = DateBegin.AddDays(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month) - 1);
			} else if (param.Equals("SetDatesToCurrentYear")) {
				DateBegin = new DateTime(DateTime.Now.Year, 1, 1);
				DateEnd = new DateTime(DateTime.Now.Year, 12, DateTime.DaysInMonth(DateTime.Now.Year, 12));
			} else if (param.Equals("GoToPreviousMonth")) {
				DateEnd = new DateTime(DateBegin.Year, DateBegin.Month, 1).AddDays(-1);
				DateBegin = DateEnd.AddDays(-1 * (DateTime.DaysInMonth(DateEnd.Year, DateEnd.Month) - 1));
			} else if (param.Equals("GoToPreviousDay")) {
				DateBegin = DateBegin.AddDays(-1);
				DateEnd = DateBegin;
			} else if (param.Equals("GoToNextDay")) {
				DateBegin = DateBegin.AddDays(1);
				DateEnd = DateBegin;
			} else if (param.Equals("GoToNextMonth")) {
				DateBegin = new DateTime(DateBegin.Year, DateBegin.Month, DateTime.DaysInMonth(DateBegin.Year, DateBegin.Month)).AddDays(1);
				DateEnd = DateBegin.AddDays((DateTime.DaysInMonth(DateBegin.Year, DateBegin.Month) - 1));
			}
		}

		private void ButtonLoadToDb_Click(object sender, RoutedEventArgs e) {
			if (MessageBox.Show(
				this,
				"Данную операцию будет невозможно отменить." + Environment.NewLine +
				"Вы уверены, что хотите записать считанные данные в БД?",
				"Запись в БД",
				MessageBoxButton.YesNo,
				MessageBoxImage.Question) == MessageBoxResult.No)
				return;

			Cursor = Cursors.Wait;
			StartWorkIndicator();
			IsButtonReadFileEnabled = false;
			IsButtonLoadToDbEnabled = false;

			bwLoadToDb = new BackgroundWorker();
			bwLoadToDb.DoWork += BwLoadToDb_DoWork;
			bwLoadToDb.RunWorkerCompleted += BwLoadToDb_RunWorkerCompleted;
			bwLoadToDb.ProgressChanged += BwReadFile_ProgressChanged;
			bwLoadToDb.WorkerReportsProgress = true;
			bwLoadToDb.RunWorkerAsync();
		}

		private void BwLoadToDb_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
			UpdateProgress(new string('=', 40));
			Cursor = Cursors.Arrow;
			IsButtonReadFileEnabled = true;
			StopWorkIndicator();

			if (e.Error != null) {
				string messageError = "Ошибка выполнения";
				Logging.ToLog(messageError);
				UpdateProgress(messageError);
				MessageBox.Show(
					this,
					e.Error.Message + Environment.NewLine + e.Error.StackTrace,
					messageError,
					MessageBoxButton.OK,
					MessageBoxImage.Error);
				return;
			}

			string message = "Завершено!";
			Logging.ToLog(message);
			UpdateProgress(message);
			MessageBox.Show(
				this,
				message,
				string.Empty,
				MessageBoxButton.OK,
				MessageBoxImage.Information);
		}

		private void BwLoadToDb_DoWork(object sender, DoWorkEventArgs e) {
			e.Result = false;
			BackgroundWorker bw = sender as BackgroundWorker;
			IsButtonLoadToDbEnabled = false;

			bw.ReportProgress(0, new string('=', 40));
			try {
				VerticaClient verticaClient = new VerticaClient(
					VerticaSettings.host,
					VerticaSettings.database,
					VerticaSettings.user,
					VerticaSettings.password,
					bw);

				bool isOk = false;
				if (IsCheckedLoadTypeTreatmentsDetails)
					isOk = verticaClient.ExecuteUpdateQuery(VerticaSettings.sqlInsertTreatmentsDetails, true, fileInfo);
				else if (IsCheckedLoadTypeProfitAndLoss)
					isOk = verticaClient.ExecuteUpdateQuery(VerticaSettings.sqlInsertProfitAndLoss, false, fileInfo);

				if (isOk) {
					bw.ReportProgress(0, "Данные успешно загружены в БД");

					if (IsCheckedLoadTypeProfitAndLoss) {
						e.Result = true;
						return;
					}

					//bw.ReportProgress(0, "Обновление данных в основной таблице с услугами (это может занять несколько минут). Дождитесь окончания.");
					//DataTable dataTableRefresh = verticaClient.GetDataTable(VerticaSettings.sqlRefreshOrderdet);
					//if (dataTableRefresh != null) 
					//	if (dataTableRefresh.Rows[0][0].ToString().Equals("refresh_columns completed")) {
					//		bw.ReportProgress(0, "Обновление выполнено успешно");

					//		bw.ReportProgress(0, "Копирование исходного файла в архив");
					//		string pathArchive = Program.GetArchivePath(out string messageError);
					//		if (string.IsNullOrEmpty(pathArchive)) {
					//			bw.ReportProgress(0, "Не удалось получить доступ к папке архива: " + messageError);
					//		} else {
					//			try {
					//				string fileName = Path.GetFileNameWithoutExtension(SelectedFile);
					//				string fileExtension = Path.GetExtension(SelectedFile);
					//				string archiveFilePath = Path.Combine(pathArchive, fileName + fileExtension);

					//				if (File.Exists(archiveFilePath))
					//					archiveFilePath = Path.Combine(pathArchive, fileName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + fileExtension);

					//				File.Copy(SelectedFile, archiveFilePath);
					//				bw.ReportProgress(0, "Файл " + fileName + fileExtension + " успешно скопирован в архив");
					//				e.Result = true;
					//			} catch (Exception excArch) {
					//				bw.ReportProgress(0, "Не удалось скопировать файл: " + excArch.Message + Environment.NewLine
					//					+ excArch.StackTrace);
					//			}
					//		}
					//	}
				} else
					bw.ReportProgress(0, "!!! Во время загрузки возникли ошибки");
			} catch (Exception exc) {
				string message = "!!! Не удалось выполнить загрузку в БД: " + 
					exc.Message + Environment.NewLine + exc.StackTrace;

				Logging.ToLog(message);
				bw.ReportProgress(0, message);
			}
		}


		private void StartWorkIndicator() {
			timerWorkIndicator.Start();
		}

		private void StopWorkIndicator() {
			timerWorkIndicator.Stop();
		}


		private void Window_Drop(object sender, DragEventArgs e) {
			if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
				string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
				string messageError = string.Empty;
				if (files.Length > 1)
					messageError = "Допускается добавляение только одного файла за раз";

					

				string file = files[0].ToLower();
				if (!(file.EndsWith(".xlsx") || file.EndsWith("*.xls")))
					messageError = "Добавить можно только книгу Excel (*.xls | *.xlsx)";

				if (!string.IsNullOrEmpty(messageError)) {
					MessageBox.Show(
						this,
						messageError,
						string.Empty,
						MessageBoxButton.OK,
						MessageBoxImage.Information);
					return;
				}

				SelectedFile = files[0];
				ReadSheetNames();
			}
		}
	}
}
