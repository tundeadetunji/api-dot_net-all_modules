# Code Repertoire

<br>
Reusable code that can be imported into .Net Framework/Core projects via <a href="https://www.nuget.org/packages/inovationware.code/#versions-body-tab" target="_blank">Nuget</a>

<p>Works for .Net Framework 4.6 and greater.</p>

<p>To use, add these imports as required (please check the individual files for up to date list):</p>

<i>C#:</i>
<pre>using static iNovation.Code.General;
using static iNovation.Code.Desktop;
using static iNovation.Code.Encryption;
using static iNovation.Code.Media;
using static iNovation.Code.PC;
using static iNovation.Code.Sequel;
using static iNovation.Code.Web;</pre>

<i>VB:</i>
<pre>
Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports iNovation.Code.Encryption
Imports iNovation.Code.Media
Imports iNovation.Code.PC
Imports iNovation.Code.Sequel
Imports iNovation.Code.Web
</pre>

<p>Then, to check if the user has typed anything in TextBox, for example:</p>

<pre>if (IsEmpty(theTextBox))
{
    //ToDo
}</pre>

<p>To encrypt text, use:</p>

<pre>String string = "to be encrypted";
String encrypted = Encrypt(string);</pre>

<p>and to get back original string:</p>

<pre>String decrypted = Decrypt(encrypted);</pre>
<br>

<h3>For non static imports:</h3>

<i>C#:</i>
<pre>
iNovation.Code.Feedback feedback = new iNovation.Code.Feedback();
</pre>
<i>VB:</i>
<pre>
Dim feedback as new iNovation.Code.Feedback()
</pre>

<p>Then to utilize Text To Speech, you can use, for example</p>

<pre>feedback.say("what to say out loud");</pre>

<br>

<h3>§</h3>
<p>
Desktop and Web  both contain the same functions but Desktop targets desktop environment while Web targets ASP.NET environment, their references differ and may conflict if used together.
</p>

<!--
<br>
<h3>Dependency:</h3>

For Feedback to work, you may need to include System.Speech version 4.0.0.0 in the same directory as the exe of the app.-->
