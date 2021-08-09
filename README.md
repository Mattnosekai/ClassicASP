# Classic ASP

There are a number of steps to take in executing a Classic ASP script. All of these steps involve configuring IIS. A few additional steps are needed to have an ASP script communicate with an ACCESS DB. 

**STEPS:**
<br>
Step #1. Confirm that IIS is running. In a web browser, type in localhost in the address bar and press enter. An IIS "welcome screen" should appear similar to the one pictured here.
<br>
![Image 1](https://github.com/Mattnosekai/ClassicASP/blob/main/asp1.png)
<br>
Step #2. Now that we have confirmed that IIS is running, a few configuration changes need to be made. The image below shows what additional boxes need to be checked.
<br>
This video also explains the same steps that need to be taken. https://www.youtube.com/watch?v=kGKCUPPk-C4
<br>
To execute this script, IIS must be enabled. Take the following steps:
<br>
Open Control Panel in Windows.
<br>
Under “Programs & Features” go to/click “Turn Windows features on or off”
<br>
Enable everything circled in red (including all sub sections):
<br>
![Image 2](https://github.com/Mattnosekai/ClassicASP/blob/main/asp2.png)
<br>
Click “OK” to save the changes. It might take a few minutes. 
![Image 3](https://github.com/Mattnosekai/ClassicASP/blob/main/asp3.png)
<br>
Next, open up a browser and type “localhost” in the address bar. This IIS screen or a similar one should be showing. This confirms that IIS is still running after the configuration changes.
<br>
![Image 1](https://github.com/Mattnosekai/ClassicASP/blob/main/asp1.png)
<br>
Next, go to c:\
<br>
Go to the inetpup directory. Inside of that, go to the wwwroot directory. 

<br>
Within that directory create a new directory called ClassicASP 
<br>
Make the source code file hello.asp
<br>
Apply the proper permissions to this file and directory if there are any issues. 
<br>
![Image 4](https://github.com/Mattnosekai/ClassicASP/blob/main/asp4.png)
<br>

In the address bar of the browser, type: localhost/ClassicASP/hello.asp 

<br>
“Hello World!” should be visible. 
<br>
<br>
![Image 5](https://github.com/Mattnosekai/ClassicASP/blob/main/asp5.png)


