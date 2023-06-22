# All Modules

<br>
Reusable code that can be imported into .Net Framework/Core projects via <a href="https://www.nuget.org/packages/inovationware.code/#versions-body-tab" target="_blank">Nuget</a>

<p>Works for .Net Framework 4.6 and greater.</p>

<p>To use, add these imports as required:</p>

<pre>using static iNovation.Code.General;
using static iNovation.Code.Desktop;
using static iNovation.Code.Encryption;
using static iNovation.Code.Bootstrap;
using static iNovation.Code.JSON;
using static iNovation.Code.Media;
using static iNovation.Code.PC;
using static iNovation.Code.Sequel;
using static iNovation.Code.Charts;
using static iNovation.Code.Web;</pre>

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

<pre>
iNovation.Code.Feedback feedback = new iNovation.Code.Feedback();
iNovation.Code.CameraAndMic capture = new iNovation.Code.CameraAndMic(DeviceCapture.SingleImage, folder_to_store_captured_files, some_System_Windows_Forms_PictureBox, number_of_seconds_to_capture_if_intended_to_be_automatic, ".jpg");
</pre>

<p>Then to utilize Text To Speech, you can use, for example</p>

<pre>feedback.say("what to say out loud");</pre>

<p>To snap a single image with webcam:</p>

<pre>capture.StartCapture();</pre>


<br>
<h3>Dependency:</h3>

For Feedback to work, you may need to include System.Speech version 4.0.0.0 in the same directory as the exe of the app.
