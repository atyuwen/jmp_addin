using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;
using System.Collections.Generic;

namespace JmpAddinVS2012
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;
			if(connectMode == ext_ConnectMode.ext_cm_UISetup)
			{
				object []contextGUIDS = new object[] { };
				Commands2 commands = (Commands2)_applicationObject.Commands;
				string editMenuName = "Edit";

				//Place the command on the tools menu.
				//Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
				Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

				//Find the Tools command bar on the MenuBar command bar:
				CommandBarControl editControl = menuBarCommandBar.Controls[editMenuName];
				CommandBarPopup editPopup = (CommandBarPopup)editControl;

				//This try/catch block can be duplicated if you wish to add multiple commands to be handled by your Add-in,
				//  just make sure you also update the QueryStatus/Exec method to include the new command names.
				try
				{
					//Add commands to the Commands collection:
                    Command commandJmpInto = commands.AddNamedCommand2(_addInInstance, "JmpInto", "Jump Into of Bracket", "Jump into the next bracket", true, 0, ref contextGUIDS,
                        (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    Command commandJmpOut = commands.AddNamedCommand2(_addInInstance, "JmpOut", "Jump Out of Bracket", "Jump out of the current bracket", true, 0, ref contextGUIDS,
                        (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    Command commandJmpOver = commands.AddNamedCommand2(_addInInstance, "JmpOver", "Jump Over of Bracket", "Jump over the adjacent bracket on the right", true, 0, ref contextGUIDS,
                        (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    // Key bindings
                    commandJmpInto.Bindings = (Object)("Text Editor::Ctrl+,");
                    commandJmpOut.Bindings = (Object)("Text Editor::Ctrl+.");
                    commandJmpOver.Bindings = (Object)("Text Editor::Tab");

					//Add a control for the command to the tools menu:
                    if ((commandJmpOut != null) && (editPopup != null))
					{
                        commandJmpOut.AddControl(editPopup.CommandBar, editPopup.Controls.Count);
					}
				}
				catch(System.ArgumentException)
				{
					//If we are here, then the exception is probably because a command with that name
					//  already exists. If so there is no need to recreate the command and we can 
                    //  safely ignore the exception.
				}
			}
		}

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		/// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
		/// <param term='commandName'>The name of the command to determine state for.</param>
		/// <param term='neededText'>Text that is needed for the command.</param>
		/// <param term='status'>The state of the command in the user interface.</param>
		/// <param term='commandText'>Text requested by the neededText parameter.</param>
		/// <seealso class='Exec' />
		public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		{
			if(neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
			{
                if (commandName == "JmpAddinVS2012.Connect.JmpInto")
				{
					status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported|vsCommandStatus.vsCommandStatusEnabled;
					return;
				}
                if (commandName == "JmpAddinVS2012.Connect.JmpOut")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }
                if (commandName == "JmpAddinVS2012.Connect.JmpOver")
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }
			}
		}

		/// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
		/// <param term='commandName'>The name of the command to execute.</param>
		/// <param term='executeOption'>Describes how the command should be run.</param>
		/// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
		/// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
		/// <param term='handled'>Informs the caller if the command was handled or not.</param>
		/// <seealso class='Exec' />
		public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		{
			handled = false;
			if(executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
			{
                if (commandName == "JmpAddinVS2012.Connect.JmpInto")
				{
					handled = JumpInto();
					return;
                } 
                if (commandName == "JmpAddinVS2012.Connect.JmpOut")
                {
                    handled = JumpOut();
                    return;
                }
                if (commandName == "JmpAddinVS2012.Connect.JmpOver")
                {
                    handled = JumpOver();
                    return;
                }
			}
		}

        /// <summary>
        /// Jump out of the nearest matched bracket
        /// </summary>
        /// <returns></returns>
        private bool JumpInto()
        {
            TextDocument doc = _applicationObject.ActiveDocument.Object("TextDocument") as TextDocument;
            EditPoint curPoint = doc.Selection.ActivePoint.CreateEditPoint();
            EditPoint afterBracketPoint = null;
            TextRanges textRangesDummy = null;

            if (InCommentLine(ref curPoint))
            {
                JumpComment(ref curPoint);
            }
            else if (InCharOrString(ref curPoint))
            {
                return false;
            }

            string regex = "[(){}\\[\\]\'\"]|(/[/*])|(\\*/)";
            char[] quotes = new char[] { '\'', '\"' };
            char[] startBrackets = new char[] { '[', '(', '{' };
            char[] endBrackets = new char[] { ']', ')', '}' };

            while (curPoint.FindPattern(regex, (int)vsFindOptions.vsFindOptionsRegularExpression,
                ref afterBracketPoint, ref textRangesDummy))
            {
                string token = curPoint.GetText(2);
                if (token == "//" || token == "/*")
                {
                    if (JumpComment(ref curPoint))
                    {
                        continue;
                    }
                    else
                    {
                        return false;
                    }
                }
                else if (token == "*/")
                {
                    return false;
                }

                char charactor = token[0];
                if (Array.IndexOf(quotes, charactor) != -1)
                {
                    break;
                }
                else if (Array.IndexOf(startBrackets, charactor) != -1)
                {
                    break;
                }
                else if (Array.IndexOf(endBrackets, charactor) != -1)
                {
                    return false;
                }
            }

            doc.Selection.MoveToPoint(curPoint, false);
            _applicationObject.ExecuteCommand("Edit.CharRight", string.Empty);
            return true;
        }

        /// <summary>
        /// Jump out of the nearest matched bracket
        /// </summary>
        /// <returns></returns>
        private bool JumpOut()
        {
            TextDocument doc = _applicationObject.ActiveDocument.Object("TextDocument") as TextDocument;
            EditPoint curPoint = doc.Selection.ActivePoint.CreateEditPoint();
            EditPoint afterBracketPoint = null;
            TextRanges textRangesDummy = null;

            if (InCommentLine(ref curPoint))
            {
                JumpComment(ref curPoint);
            }
            else if (InCharOrString(ref curPoint))
            {
                JumpCharOrString(ref curPoint);
                doc.Selection.MoveToPoint(curPoint, false);
                return true;
            }

            Stack<char> brackets = new Stack<char>();
            string regex = "[(){}\\[\\]\'\"]|(/[/*])|(\\*/)";
            char[] quotes = new char[] { '\'', '\"' };
            char[] startBrackets = new char[] { '[', '(', '{' };
            char[] endBrackets = new char[] { ']', ')', '}' };

            while (curPoint.FindPattern(regex, (int)vsFindOptions.vsFindOptionsRegularExpression,
                ref afterBracketPoint, ref textRangesDummy))
            {
                string token = curPoint.GetText(2);
                if (token == "//" || token == "/*")
                {
                    if (JumpComment(ref curPoint))
                    {
                        continue;
                    }
                    else
                    {
                        return false;
                    }
                }
                else if (token == "*/")
                {
                    curPoint.CharRight(1);
                    break;
                }

                char charactor = token[0];
                if (Array.IndexOf(quotes, charactor) != -1)
                {
                    if (JumpCharOrString(ref curPoint))
                    {
                        continue;
                    }
                    else
                    {
                        return false;
                    }
                }
                else if (Array.IndexOf(startBrackets, charactor) != -1)
                {
                    brackets.Push(charactor);
                    curPoint.CharRight(1);
                }
                else
                {
                    int index = Array.IndexOf(endBrackets, charactor);

                    if (brackets.Count == 0)
                    {
                        break;
                    }
                    else if (brackets.Pop() == startBrackets[index])
                    {
                        curPoint.CharRight(1);
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            doc.Selection.MoveToPoint(curPoint, false);
            _applicationObject.ExecuteCommand("Edit.CharRight", string.Empty);
            return true;
        }

        /// <summary>
        /// Jump over the adjacent bracket on the right
        /// </summary>
        /// <returns></returns>
        private bool JumpOver()
        {
            TextDocument doc = _applicationObject.ActiveDocument.Object("TextDocument") as TextDocument;
            EditPoint curPoint = doc.Selection.ActivePoint.CreateEditPoint();

            char charactor = curPoint.GetText(1)[0];
            if (charactor == ' ' || charactor == '\t')
            {
                curPoint.CharRight(1);
                charactor = curPoint.GetText(1)[0];
            }

            char[] brackets = new char[] { '\'', '\"', '[', ']', '(', ')', '{', '}', '<', '>', ',', ';' };
            if (Array.IndexOf(brackets, charactor) != -1)
            {
                EditPoint headPoint = curPoint.CreateEditPoint();
                headPoint.StartOfLine();
                string line = headPoint.GetText(curPoint);

                if (line.TrimStart() != string.Empty)
                {
                    doc.Selection.MoveToPoint(curPoint, false);
                    _applicationObject.ExecuteCommand("Edit.CharRight", string.Empty);
                    return true;
                }
            }

            _applicationObject.ExecuteCommand("Edit.InsertTab", string.Empty);
            return false;
        }

        /// <summary>
        /// Judge whether the current edit point is in a char or string
        /// </summary>
        /// <param name="point"></param>
        /// <returns></returns>
        private bool InCharOrString(ref EditPoint point)
        {
            EditPoint newPoint = point.CreateEditPoint();
            newPoint.StartOfLine();
            int prevPos = newPoint.AbsoluteCharOffset;

            while (newPoint.LessThan(point))
            {
                char charactor = newPoint.GetText(1)[0];
                if (charactor == '\'' || charactor == '\"')
                {
                    prevPos = newPoint.AbsoluteCharOffset;
                    JumpCharOrString(ref newPoint);
                }
                else
                {
                    newPoint.CharRight(1);
                }
            }

            if (point.LessThan(newPoint))
            {
                point.MoveToAbsoluteOffset(prevPos);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Jump over a char or string
        /// </summary>
        /// <param name="point"></param>
        /// <returns></returns>
        private bool JumpCharOrString(ref EditPoint point)
        {
            char literal = point.GetText(1)[0];
            point.CharRight(1);

            while (true)
            {
                char charactor = point.GetText(1)[0];
                point.CharRight(1);

                if (charactor == literal)
                {
                    return true;
                }

                if (charactor == '\n')
                {
                    return false;
                }

                if (charactor == '\\')
                {
                    point.CharRight(1);
                    continue;
                }
            }
        }

        /// <summary>
        /// Judge whether the current edit point is in a comment line
        /// </summary>
        /// <param name="point"></param>
        /// <returns></returns>
        private bool InCommentLine(ref EditPoint point)
        {
            EditPoint newPoint = point.CreateEditPoint();
            newPoint.StartOfLine();

            while (newPoint.LessThan(point))
            {
                string token = newPoint.GetText(2);
                if (token == "//")
                {
                    point.MoveToPoint(newPoint);
                    return true;
                }

                char charactor = token[0];
                if (charactor == '\'' || charactor == '\"')
                {
                    JumpCharOrString(ref newPoint);
                }
                else
                {
                    newPoint.CharRight(1);
                }
            }
            return false;
        }

        /// <summary>
        /// Jump over a comment block
        /// </summary>
        /// <param name="point"></param>
        /// <returns></returns>
        private bool JumpComment(ref EditPoint point)
        {
            string token = point.GetText(2);

            if (token == "//")
            {
                point.EndOfLine();
                if (!point.AtEndOfDocument)
                {
                    point.LineDown(1);
                    point.StartOfLine();
                }
                return true;
            }
            else // if (token == "/*")
            {
                EditPoint afterBracketPoint = null;
                TextRanges textRangesDummy = null;
                if (point.FindPattern("*/", 0, ref afterBracketPoint, ref textRangesDummy))
                {
                    point.CharRight(2);
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

		private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}