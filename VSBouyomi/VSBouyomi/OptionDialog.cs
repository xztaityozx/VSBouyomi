using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;

namespace VSBouyomi {
	public class OptionDialog : DialogPage
	{
		private const string Fixedtext = @"発言する内容です";

		[Category("発言内容設定")]
		[DisplayName("ビルド成功時")]
		[Description("ビルド成功時に"+Fixedtext+"{ProjectName}と書くとプロジェクト名、{Platfrom}と書くとプラットフォーム名に変わります")]
		public TalkData BuildDoneTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("ビルド失敗時")]
		[Description("ビルド失敗時に"+Fixedtext+"{ProjectName}と書くとプロジェクト名、{Platfrom}と書くとプラットフォーム名に変わります")]
		public TalkData BuildFaildTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("ビルド開始時")]
		[Description("ビルド開始時に"+Fixedtext+"{ProjectName}と書くとプロジェクト名、{Platfrom}と書くとプラットフォーム名に変わります")]
		public TalkData BuildBeginTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("デザインモードになったとき")]
		[Description("デザインモードになったときに"+Fixedtext)]
		public TalkData EnterDesignTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("デバッガーモードになったとき")]
		[Description("デバッガーモードになったときに" + Fixedtext)]
		public TalkData EnterDebuggerTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("例外がスローされたとき")]
		[Description("例外がスローされたときに" + Fixedtext)]
		public TalkData ExceptionThrownTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("ハンドルされていない例外がスローされたとき")]
		[Description("ハンドルされていない例外がスローされたときに" + Fixedtext)]
		public TalkData ExceptionNotHandledTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("VisualStudioが起動したとき")]
		[Description("VisualStudioが起動したときに"+Fixedtext)]
		public TalkData StartupCompleteTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("VisualStudioが終了するとき")]
		[Description("VisualStudioが終了するとき"+Fixedtext)]
		public TalkData BeginShutdownTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("プロジェクトを開いたあと")]
		[Description("プロジェクトを開いたあと" + Fixedtext)]
		public TalkData OpenedTalkData { get; set; }

		[Category("発言内容設定")]
		[Description("デバッガーで実行モードになったときに"+Fixedtext)]
		[DisplayName("デバッガーで実行モードになったとき")]
		public TalkData EnterRunModeTalkData { get; set; }

		[Category("発言内容設定")]
		[DisplayName("プロジェクトを閉じた後")]
		[Description("プロジェクトを閉じた後" + Fixedtext)]
		public TalkData ClosingTalkData { get; set; }

		[Category("ネットワーク設定")]
		[DisplayName("IPアドレス")]
		[Description("棒読みちゃんのIPアドレスです")]
		public string IPAddress { get; set; }

		[Category("ネットワーク設定")]
		[DisplayName("ポート番号")]
		[Description("棒読みちゃんのポート番号です")]
		public int Port { get; set; }


		public override void LoadSettingsFromStorage() {
			var config = Eventer.ReadConfigData(ActivatePackage.ConfigXmlPath);
			BuildDoneTalkData = config.BuildDone;
			Port = config.Port;
			IPAddress = config.IPAddress;
			BuildFaildTalkData = config.BuildFaild;
			BuildBeginTalkData = config.BuildBegin;
			StartupCompleteTalkData = config.StartupComplete;
			BeginShutdownTalkData = config.BeginShutdown;
			OpenedTalkData = config.Opened;
			ClosingTalkData = config.Closing;
			ExceptionNotHandledTalkData = config.ExceptionNotHandled;
			ExceptionThrownTalkData = config.ExceptionThrown;
			EnterDebuggerTalkData = config.EnterDebugger;
			EnterDesignTalkData = config.EnterDesign;
			EnterRunModeTalkData = config.EnterRunMode;
		}

		public override void SaveSettingsToStorage() {
			var config = new VsBouyomiConfigData
			{
				BuildDone = BuildDoneTalkData,
				BuildFaild = BuildFaildTalkData,
				BuildBegin = BuildBeginTalkData,
				StartupComplete = StartupCompleteTalkData,
				BeginShutdown = BeginShutdownTalkData,
				Opened = OpenedTalkData,
				Closing = ClosingTalkData,
				ExceptionNotHandled = ExceptionNotHandledTalkData,
				ExceptionThrown = ExceptionThrownTalkData,
				EnterDebugger = EnterDebuggerTalkData,
				EnterDesign = EnterDesignTalkData,
				EnterRunMode = EnterRunModeTalkData,
				IPAddress = IPAddress,
				Port = Port
			};
			Eventer.WriteConfigData(ActivatePackage.ConfigXmlPath, config);
			Activate.Eventer.Config = config;
		}
	}
}
