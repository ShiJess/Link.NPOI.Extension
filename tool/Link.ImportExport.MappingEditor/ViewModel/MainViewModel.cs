using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.Win32;
using System.IO;
using System.Windows.Input;

namespace Link.ImportExport.MappingEditor.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// Use the <strong>mvvminpc</strong> snippet to add bindable properties to this ViewModel.
    /// </para>
    /// <para>
    /// You can also use Blend to data bind with the tool's support.
    /// </para>
    /// <para>
    /// See http://www.galasoft.ch/mvvm
    /// </para>
    /// </summary>
    public class MainViewModel : ViewModelBase
    {

        private MappingConfig config = new MappingConfig();
        /// <summary>
        /// �༭������ģ��
        /// </summary>
        public MappingConfig Config
        {
            get { return config; }
            set { Set(() => Config, ref config, value); }
        }

        private string fullfilename = string.Empty;
        /// <summary>
        /// ģ������·��
        /// </summary>
        public string FullFileName
        {
            get { return fullfilename; }
            set { Set(() => FullFileName, ref fullfilename, value); }
        }

        /// <summary>
        /// ��ģ���ļ�����
        /// </summary>
        public ICommand OpenMappingFileCommand { get; set; }
        /// <summary>
        /// ����ģ���ļ�����
        /// </summary>
        public ICommand SaveMappingFileCommand { get; set; }

        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        public MainViewModel()
        {
            if (IsInDesignMode)
            {
                // Code runs in Blend --> create design time data.
            }
            else
            {
                // Code runs "for real"
                OpenMappingFileCommand = new RelayCommand(OpenMappingFile);
                SaveMappingFileCommand = new RelayCommand(SaveMappingFile);
            }
        }

        /// <summary>
        /// ��ģ���ļ�
        /// </summary>
        public void OpenMappingFile()
        {
            FileDialog fd = new OpenFileDialog();
            fd.Filter = "*.xml|*.xml";
            if (fd.ShowDialog() ?? false)
            {
                FullFileName = fd.FileName;
                Config = MappingConfig.ReadFromXmlFormat(FullFileName);
            }
        }

        /// <summary>
        /// ����ģ���ļ�
        /// </summary>
        public void SaveMappingFile()
        {
            if (string.IsNullOrWhiteSpace(FullFileName) || !File.Exists(FullFileName))
            {
                FileDialog fd = new SaveFileDialog();
                fd.Filter = "*.xml|*.xml";
                fd.FileName = "�½�ģ��.xml";
                if (fd.ShowDialog() ?? false)
                {
                    FullFileName = fd.FileName;
                }
                else
                {
                    return;
                }
            }
            MappingConfig.SaveAsXmlFormat(FullFileName, Config);
        }

    }
}