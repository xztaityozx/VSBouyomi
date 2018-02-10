using EnvDTE;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace VSBouyomi {
	[TypeConverter(typeof(TalkDataConverter))]
	public class TalkData
	{
		[DefaultValue("")]
		[NotifyParentProperty(true)]
		[DisplayName("テキスト")]
		public string Text { get; set; }
		[DefaultValue(false)]
		[NotifyParentProperty(true)]
		[DisplayName("すぐに？")]
		[Description("イベントが発生したらすぐに読み上げるかどうかです")]
		public bool Immediate { get; set; }
	}
	public class VsBouyomiConfigData {
		public TalkData BuildDone { get; set; }
		public TalkData BuildFaild { get; set; }
		public TalkData BuildBegin { get; set; }
		public TalkData StartupComplete { get; set; }
		public TalkData BeginShutdown { get; set; }
		public TalkData Opened { get; set; }
		public TalkData Closing { get; set; }
		public TalkData ExceptionNotHandled { get; set; }
		public TalkData ExceptionThrown { get; set; }
		public TalkData EnterDebugger { get; set; }
		public TalkData EnterDesign { get; set; }
		public TalkData EnterRunMode { get; set; }
		public string IPAddress { get; set; }
		public int Port { get; set; }
	}

	public class TalkDataConverter : ExpandableObjectConverter {
		public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType) {
			return sourceType == typeof(string) || base.CanConvertFrom(context, sourceType);
		}

		public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType) {
			return destinationType == typeof(string) || base.CanConvertTo(context, destinationType);
		}

		public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value) {
			if (!(value is string str)) return base.ConvertFrom(context, culture, value);
			var values = str.Split(',');
			var length = values.Length;
			try {
				return new TalkData
				{
					Text = (length < 1 || values[0].Trim().Length == 0) ? "" : values[0],
					Immediate = (length >= 2 && values[1].Trim().Length != 0) &&(bool) new BooleanConverter().ConvertFromString(values[1])
				};
			}
			catch (Exception) {
				throw new ArgumentException("プロパティの値が無効です");
			}	
		}

		public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType) {
			if (!(value is TalkData td)||destinationType!=typeof(string)) return base.ConvertTo(context, culture, value,destinationType);
			return $"{td.Text},{td.Immediate.ToString()}";
		}
	}

	public class Eventer
	{
		private readonly EventHolder _holder;
		private BouyomiTcpClient BouyomiTcpClient { get; set; }
		public VsBouyomiConfigData Config { get; set; }

		public Eventer(EventHolder h,string path) {
			_holder = h;
			Config = ReadConfigData(path);
			BouyomiTcpClient = new BouyomiTcpClient(Config.IPAddress, Config.Port);
		}

		public void Activate() {
			_holder.BuildEvents.OnBuildDone += BuildEvents_OnBuildDone;
			_holder.BuildEvents.OnBuildBegin += BuildEvents_OnBuildBegin;
			_holder.BuildEvents.OnBuildProjConfigDone += BuildEvents_OnBuildProjConfigDone;
			_holder.DebuggerEvents.OnEnterBreakMode += DebuggerEvents_OnEnterBreakMode;
			_holder.DebuggerEvents.OnEnterRunMode += DebuggerEvents_OnEnterRunMode;
			_holder.DteEvents.ModeChanged += DteEvents_ModeChanged;
			_holder.DteEvents.OnBeginShutdown += DteEvents_OnBeginShutdown;
			_holder.DteEvents.OnStartupComplete += DteEvents_OnStartupComplete;
			_holder.SolutionEvents.Opened += SolutionEvents_Opened;
			_holder.SolutionEvents.AfterClosing += SolutionEvents_AfterClosing;
			_holder.BuildEvents.OnBuildProjConfigBegin += BuildEvents_OnBuildProjConfigBegin;
		}



		public void Deactivate() {
			_holder.BuildEvents.OnBuildDone -= BuildEvents_OnBuildDone;
			_holder.BuildEvents.OnBuildBegin -= BuildEvents_OnBuildBegin;
			_holder.BuildEvents.OnBuildProjConfigDone -= BuildEvents_OnBuildProjConfigDone;
			_holder.DebuggerEvents.OnEnterBreakMode -= DebuggerEvents_OnEnterBreakMode;
			_holder.DebuggerEvents.OnEnterRunMode -= DebuggerEvents_OnEnterRunMode;
			_holder.DteEvents.ModeChanged -= DteEvents_ModeChanged;
			_holder.DteEvents.OnBeginShutdown -= DteEvents_OnBeginShutdown;
			_holder.DteEvents.OnStartupComplete -= DteEvents_OnStartupComplete;
			_holder.SolutionEvents.Opened -= SolutionEvents_Opened;
			_holder.SolutionEvents.AfterClosing -= SolutionEvents_AfterClosing;
			_holder.BuildEvents.OnBuildProjConfigBegin -= BuildEvents_OnBuildProjConfigBegin;
		}

		public static VsBouyomiConfigData ReadConfigData(string path) {
			var serializer = new XmlSerializer(typeof(VsBouyomiConfigData));
			using (var sr = new StreamReader(path,Encoding.UTF8)) {
				return serializer.Deserialize(sr) as VsBouyomiConfigData;
			}
		}

		public static void WriteConfigData(string path, VsBouyomiConfigData config) {
			var serializer = new XmlSerializer(typeof(VsBouyomiConfigData));
			using (var sw = new StreamWriter(path, false, Encoding.UTF8)) {
				serializer.Serialize(sw, config);
			}
		}

		private void SendTalk(TalkData td) {
			if(td.Text=="")return;
			BouyomiTcpClient.Talk(td.Text, td.Immediate);
		}

		private void SolutionEvents_AfterClosing() {
			SendTalk(Config.Closing);
		}

		private void SolutionEvents_Opened() {
			SendTalk(Config.Opened);
		}

		private void DteEvents_OnStartupComplete() {
			SendTalk(Config.StartupComplete);
		}

		private void DteEvents_OnBeginShutdown() {
			SendTalk(Config.BeginShutdown);
		}

		private void DteEvents_ModeChanged(vsIDEMode LastMode) {
			switch (LastMode) {
				case vsIDEMode.vsIDEModeDesign:
					SendTalk(Config.EnterDebugger);
					break;
				case vsIDEMode.vsIDEModeDebug:
					SendTalk(Config.EnterDesign);
					break;
				default:
					break;
			}
		}

		private void DebuggerEvents_OnEnterRunMode(dbgEventReason Reason) {
			SendTalk(Config.EnterRunMode);
		}

		private void DebuggerEvents_OnEnterBreakMode(dbgEventReason Reason, ref dbgExecutionAction ExecutionAction) {
			if (Reason == dbgEventReason.dbgEventReasonExceptionNotHandled) SendTalk(Config.ExceptionNotHandled);
			else if (Reason == dbgEventReason.dbgEventReasonExceptionThrown) SendTalk(Config.ExceptionThrown);
		}

		private readonly Dictionary<string, string> dictionary = new Dictionary<string, string>
		{
			{"{ProjectName}", ""},
			{"{Platform}", ""},
		};

		private bool lastCondition = true;
		private void BuildEvents_OnBuildProjConfigDone(string Project, string ProjectConfig, string Platform, string SolutionConfig, bool Success) {
			lastCondition = Success;
			dictionary["{ProjectName}"] = Project;
			dictionary["{Platform}"] = Platform;
			Debug.WriteLine($"{nameof(ProjectConfig)}:{ProjectConfig}\n{nameof(SolutionConfig)}:{SolutionConfig}");
		}

		private void BuildEvents_OnBuildBegin(vsBuildScope Scope, vsBuildAction Action) {
			SendTalk(Config.BuildBegin);
		}

		private void BuildEvents_OnBuildDone(vsBuildScope Scope, vsBuildAction Action) {
			SendTalk(BuildEventFilter(lastCondition ? Config.BuildDone : Config.BuildFaild));
		}

		private TalkData BuildEventFilter(TalkData td) {
			var text = dictionary.Aggregate(td.Text, (current, item) => current.Replace(item.Key, item.Value));
			return new TalkData() {Immediate = td.Immediate, Text = text};
		}
		private void BuildEvents_OnBuildProjConfigBegin(string Project, string ProjectConfig, string Platform, string SolutionConfig) {
			dictionary["{ProjectName}"] = Project;
			dictionary["{Platform}"] = Platform;
		}
	}
}
