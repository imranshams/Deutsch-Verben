// Copyright © 2010-2015 The CefSharp Authors. All rights reserved.
//
// Use of this source code is governed by a BSD-style license that can be found in the LICENSE file.

using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using CefSharp.MinimalExample.WinForms.Controls;
using CefSharp.MinimalExample.WinForms.CustomeJs;
using CefSharp.WinForms;
using HtmlAgilityPack;
using System.Web;

namespace CefSharp.MinimalExample.WinForms
{
    public partial class BrowserForm : Form
    {
        private readonly ChromiumWebBrowser browser;
        private int lastIndex = -1;
        private string lastVerb = "";

        private List<ExcelRecord> allVerbs = new List<ExcelRecord>();

        public string FilePath { get { return Environment.CurrentDirectory + @"\..\..\..\Data\B1.xlsx"; } }

        public BrowserForm()
        {
            InitializeComponent();

            Text = "CefSharp";
            WindowState = FormWindowState.Maximized;

            allVerbs = Excel.ReadFile(FilePath);


            browser = new ChromiumWebBrowser("https://deutsch.lingolia.com/en/grammar/conjugator")
            {
                Dock = DockStyle.Fill,
            };
            toolStripContainer.ContentPanel.Controls.Add(browser);

            browser.LoadingStateChanged += OnLoadingStateChanged;
            browser.ConsoleMessage += OnBrowserConsoleMessage;
            browser.StatusMessage += OnBrowserStatusMessage;
            browser.TitleChanged += OnBrowserTitleChanged;
            browser.AddressChanged += OnBrowserAddressChanged;
            browser.FrameLoadEnd += Browser_FrameLoadEnd;

            browser.RegisterJsObject("jsObj", new jsObj(browser));

            var bitness = Environment.Is64BitProcess ? "x64" : "x86";
            var version = String.Format("Chromium: {0}, CEF: {1}, CefSharp: {2}, Environment: {3}", Cef.ChromiumVersion, Cef.CefVersion, Cef.CefSharpVersion, bitness);
            DisplayOutput(version);
        }

        private void Browser_FrameLoadEnd(object sender, FrameLoadEndEventArgs e)
        {
            if (e.HttpStatusCode == 200 && e.Frame.IsMain)
            {
                browser.GetSourceAsync().ContinueWith(taskHtml =>
                {
                    var html = taskHtml.Result;
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(html);

                    var filterType = doc.DocumentNode.Descendants("div").Where(div => div.HasClass("container_menus") && div.HasClass("border_text")).FirstOrDefault();
                    var filterVerb = doc.DocumentNode.Descendants("input").Where(input => input.Attributes.Where(att => att.Name == "name" && att.Value == "search_verb[input]").Count() > 0).FirstOrDefault()?.GetAttributeValue("value", "");

                    //allVerbs[candidateIndex].Präsens = "updated";

                    //This is the first time
                    if (string.IsNullOrEmpty(filterVerb))
                    {
                        FindNextRecord();
                    }
                    else
                    {
                        if (filterVerb != lastVerb)
                        {
                            allVerbs[lastIndex].Flag = true;
                        }

                        var indikativDiv = doc.DocumentNode
                            .Descendants("div")
                            .Where(div =>
                                div.HasClass("container_tables") &&
                                div.HasClass("hiddeable") &&
                                div.HasClass("container_default") &&
                                div.HasClass("active")).FirstOrDefault();

                        var infiniteFormen = doc.DocumentNode
                            .Descendants("div")
                            .Where(div =>
                                div.HasClass("container_tables") &&
                                div.HasClass("active")).FirstOrDefault();

                        if (indikativDiv != null && infiniteFormen != null)
                        {
                            var divs = indikativDiv.Descendants("div").Where(div => div.HasClass("table_wrapper")).ToList();
                            var textPresent = FormatText(divs[0]);
                            var textPreterium = FormatText(divs[1]);
                            var textPerfeckt = FormatText(divs[2]);
                            var partizip = infiniteFormen.Descendants("tr").ToList()[2].Descendants("td").ToList()[1].InnerText;
                            var infinitive = infiniteFormen.Descendants("tr").ToList()[1].Descendants("td").ToList()[1].InnerText;
                            var bewegung = FindBewegung(divs[2], partizip);

                            ////multiple verb
                            //if (!string.IsNullOrEmpty(filterType.InnerHtml))
                            //{
                            //    foreach (var type in filterType.Descendants("li").ToList())
                            //    {
                            //        switch (type.InnerText.Trim().ToLower())
                            //        {
                            //            case "irregular":
                            //            case "regular":
                            //            case "not reflexive":
                            //                {
                            //                    UpdateRecord(infinitive, textPresent, textPreterium);
                            //                    break;
                            //                }
                            //            case "reflexive":
                            //                {
                            //                    UpdateRecord(infinitive, textPresent, textPreterium);
                            //                    break;
                            //                }
                            //            default:
                            //                break;
                            //        }
                            //    }
                            //}
                            //else
                            //{
                            //    UpdateRecord(infinitive, textPresent, textPreterium);
                            //}

                            UpdateRecord(infinitive, textPresent, textPreterium, partizip, bewegung);

                            FindNextRecord();
                        }
                        else
                        {
                            allVerbs[lastIndex].Flag = true;
                            FindNextRecord();
                        }
                    }
                });
                //browser.ExecuteScriptAsync("jsObj.showDevTools();");
            }
        }

        private void UpdateRecord(string filterVerb, string present, string preterium, string partizip, string bewegung)
        {
            var verb = allVerbs.Where(x => x.FinalVerb == filterVerb).SingleOrDefault();
            if (verb != null)
            {
                verb.Präsens = present;
                verb.Präteritum = preterium;
                verb.Partizip = partizip;
                verb.Perfekt = bewegung;
                Excel.UpdateRecord(FilePath, verb);
            }
            else
            {
                allVerbs.Remove(allVerbs.Where(x => x.FinalVerb == filterVerb || x.FinalVerb == "sich " + filterVerb).SingleOrDefault());
            }
        }

        private void FindNextRecord()
        {
            var candidateIndex = allVerbs.FindIndex(x => (string.IsNullOrEmpty(x.Präsens) || string.IsNullOrEmpty(x.Partizip) || string.IsNullOrEmpty(x.Perfekt)) && !x.Flag);
            if (candidateIndex >= 0)
            {
                var verb = allVerbs[candidateIndex].FinalVerb;
                lastIndex = candidateIndex;
                SearchVerb(verb);
            }
            else
            {
                Close();
            }
        }

        private string FormatText(HtmlNode node)
        {
            var result = "";
            var tds = node.Descendants("td").Where(div => div.HasClass("lia_verb")).ToList();
            foreach (var td in tds)
            {
                var value = HttpUtility.HtmlDecode(td.InnerText).Trim();
                if (!string.IsNullOrEmpty(value))
                {
                    result += value + ", ";
                }
            }
            return result.Remove(result.Length - 2, 2);
        }

        private string FindBewegung(HtmlNode node, string partizip)
        {
            var result = HttpUtility.HtmlDecode(node.Descendants("tbody").First().Descendants("tr").Last().Descendants("td").Last().InnerText).Replace(partizip, "").Trim();
            if (result == "sind")
            {
                result = "sein";
            }
            return result;
        }

        private void SearchVerb(string verb)
        {
            lastVerb = verb;
            browser.ExecuteScriptAsync("document.getElementById('search_verb_input').value = '" + verb + "';");
            browser.ExecuteScriptAsync("document.getElementById('conjugator_form').submit();");
        }

        private void OnBrowserConsoleMessage(object sender, ConsoleMessageEventArgs args)
        {
            DisplayOutput(string.Format("Line: {0}, Source: {1}, Message: {2}", args.Line, args.Source, args.Message));
        }

        private void OnBrowserStatusMessage(object sender, StatusMessageEventArgs args)
        {
            this.InvokeOnUiThreadIfRequired(() => statusLabel.Text = args.Value);
        }

        private void OnLoadingStateChanged(object sender, LoadingStateChangedEventArgs args)
        {
            SetCanGoBack(args.CanGoBack);
            SetCanGoForward(args.CanGoForward);

            this.InvokeOnUiThreadIfRequired(() => SetIsLoading(!args.CanReload));
        }

        private void OnBrowserTitleChanged(object sender, TitleChangedEventArgs args)
        {
            this.InvokeOnUiThreadIfRequired(() => Text = args.Title);
        }

        private void OnBrowserAddressChanged(object sender, AddressChangedEventArgs args)
        {
            this.InvokeOnUiThreadIfRequired(() => urlTextBox.Text = args.Address);
        }

        private void SetCanGoBack(bool canGoBack)
        {
            this.InvokeOnUiThreadIfRequired(() => backButton.Enabled = canGoBack);
        }

        private void SetCanGoForward(bool canGoForward)
        {
            this.InvokeOnUiThreadIfRequired(() => forwardButton.Enabled = canGoForward);
        }

        private void SetIsLoading(bool isLoading)
        {
            goButton.Text = isLoading ?
                "Stop" :
                "Go";
            goButton.Image = isLoading ?
                Properties.Resources.nav_plain_red :
                Properties.Resources.nav_plain_green;

            HandleToolStripLayout();
        }

        public void DisplayOutput(string output)
        {
            this.InvokeOnUiThreadIfRequired(() => outputLabel.Text = output);
        }

        private void HandleToolStripLayout(object sender, LayoutEventArgs e)
        {
            HandleToolStripLayout();
        }

        private void HandleToolStripLayout()
        {
            var width = toolStrip1.Width;
            foreach (ToolStripItem item in toolStrip1.Items)
            {
                if (item != urlTextBox)
                {
                    width -= item.Width - item.Margin.Horizontal;
                }
            }
            urlTextBox.Width = Math.Max(0, width - urlTextBox.Margin.Horizontal - 18);
        }

        private void ExitMenuItemClick(object sender, EventArgs e)
        {
            browser.Dispose();
            Cef.Shutdown();
            Close();
        }

        private void GoButtonClick(object sender, EventArgs e)
        {
            LoadUrl(urlTextBox.Text);
        }

        private void BackButtonClick(object sender, EventArgs e)
        {
            browser.Back();
        }

        private void ForwardButtonClick(object sender, EventArgs e)
        {
            browser.Forward();
        }

        private void UrlTextBoxKeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
            {
                return;
            }

            LoadUrl(urlTextBox.Text);
        }

        private void LoadUrl(string url)
        {
            if (Uri.IsWellFormedUriString(url, UriKind.RelativeOrAbsolute))
            {
                browser.Load(url);
            }
        }
    }
}
