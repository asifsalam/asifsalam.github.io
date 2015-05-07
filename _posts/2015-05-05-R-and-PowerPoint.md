---
layout: post
title: PowerPoint from R - How to automate PowerPoint with R
---

For many of us, there's just no getting away from PowerPoint.  So if you use R (or Python) for data analysis, and PowerPoint for presenting/distributing the results, it's worthwhile learning how to include this last step into your R workflow.  In this talk, [S Anand uses Python and **pywin32**][1] to do some pretty impressive things with Powerpoint, from precisely placing images, to some eye-popping animation.    

[**RDCOMClient**][2] by Duncan Temple Lang allows you to do the same thing using R. It provides the ability to "access and control applications such as Excel, Word, Power Point, Web browsers etc."  Beyond efficiency and repeatability, programmatic access enables you to do things that just would not be possible with point and click.  

<p> </p>
<!-- <div class="warning" style="height:50px;border:1px solid #d9d9d9; background-color:#f2eeee; text-align:left; vertical-align:middle; "> -->
<div class="warning" style="background-color:#f2eeee;"><p>
<img class="centre_image" src="/images/caution_finland_road_sign_189.svg" alt="Caution" style="width:50px; margin:2px 0 0 5px; align:left;" ><b style="margin: 0 0 0 10px"><i>Gratuitous animation ahead</i></b></p>
</div>

In this post, we'll go through how to recreate one of the slides from Anand's talk.  We'll scrape some data on Clint Eastwood from IMDB In the end, we should have a pretty neat PowerPoint slide with Clint Eastwood's filmography, using data from [IMDB][4].
(Clearly PowerPoint has a powerful animation engine, and the sophisticated object model allows you to programmatically manipulate almost everything.  It would have been too easy with good documentation too, so Microsoft has organized it like a puzzle with pieces hidden all over the place.)

Let's get started.
##Setup
I'm using:
  * Windows 7 (64bit)
* Office 2013 (64bit)
* R version 3.1.3
* RStudio verison 0.98

We'll need a few packages.  
1. [RDOCMClient][2] by Duncan Temple Lang. COM client that allows us to manipulate PowerPoint through R.  
2. [XML][3] by Duncal Temp Lang. Parse XML files.  
3. [rvest][4] by Hadley Wickham. Easily scrape data from the web.

Load the packages:
```{r}
install.packages(c("RDCOMClient","XML","rvest"))  
library("RDCOMClient")  
library("XML")  
library("rvest")
```

##RDCOMClient Basics
Some RDCOM basics we'll need to know:  
* `COMCreate` starts a new instance of an application.  
* Executing methods takes the form `comObj$methodName(arg1,arg2,arg3,...)`.  
* Setting properties takes the form `myObj[["Property"]] = TRUE`.  

With these basic capabilities, we can access and manipulate all the objects exposed in the PowerPoint API.

Let's create a new PowerPoint presentation from R
In R


  



  
<!-- ![Caution]({{ site.baseurl }}/images/caution_finland_road_sign_189.svg) -->


[1]:https://www.youtube.com/watch?v=aKCXj1DyEhM "S Anand"
[2]:http://www.omegahat.org/RDCOMClient/ "RCDOMClient"
[3]:http://www.omegahat.org/RSXML/ "XML Package for R"
[4]:http://www.imdb.org/ "IMDB"
[5]:https://msdn.microsoft.com/en-us/library/office/ee861525.aspx "PowerPoint 2013 Developer Reference"
[6]:https://msdn.microsoft.com/en-us/library/office/gg190747(v=office.14).aspx
[7]:https://msdn.microsoft.com/en-us/library/office/aa211626(v=office.11).aspx