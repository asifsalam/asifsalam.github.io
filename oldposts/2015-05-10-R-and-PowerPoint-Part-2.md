---
layout: post
title: How to Create Amazing PowerPoint Slides using R - Part 2 (3)
---

Now that we have a few basic tools for manipulating PowerPoint slides, let's use the power of R to build really amazing slides. In this post we'll get some data that we will use in Part 3 for constructing a PowerPoint slide with lots of (needless) animation, and some interaction.

You should have the same R/RStudio setup and the packages we loaded in [Part 1][1]

Let's get started.









Let's face it, PowerPoint isn't going anywhere. Even if you use R (or Python) for data analysis, PowerPoint is how you distribute and communicate results, and learning how to create those decks as part of your R workflow can pay off.  Beyond efficiency, and repeatability, programmatic access enables you to do things that just aren't possible with point and click.  In this tutorial, we'll learn how to automate PowerPoint using R.

In this talk, [S Anand uses Python and **pywin32**][1] to create some jaw dropping effects in Powerpoint, scraping data from IMDB and creating a PowerPoint slide using the data.  [**RDCOMClient**][2] by Duncan Temple Lang allows you to do the same thing using R. It provides the ability to "access and control applications such as Excel, Word, Power Point, Web browsers etc."

We will recreate some elements of [S Anand's talk][1], and a few other things, with a focus on interaction and animation. We'll learn the basics of accessing methods and properties of PowerPoint VBA objects, scrape data on Clint Eastwood's movies from [IMDB][13], and use it to create a slide.  In the end, we should have a cool, fun PowerPoint slide with Clint Eastwood's filmography.  

<div class="warning"><p style="margin: 0 0 0 10px">
    <img class="centre_image" src="/images/caution_finland_road_sign_189.svg" alt="Caution" style="width:50px; margin:2px 0 0 5px; align:left;" >
    <b><i>Gratuitous animation ahead</i></b></p>
</div>

(Clearly PowerPoint has a powerful animation engine, and the [object model][6] allows you to programmatically manipulate almost everything. The [Microsoft documentation][5], however, seems to be organized as a challenge, at least for beginners.)

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
3. [rvest][4] by Hadley Wickham. Web scraping simplified.

Start RStudio and Load the packages:

```r  
install.packages(c("RDCOMClient","XML","rvest"))  
library("RDCOMClient")   
library("XML")   
library("rvest")   
```

##RDCOMClient Basics
Some RDCOM basics we'll need to know: 

* To start a new instance of an application:  `COMCreate("xxx.Application")`
* Executing methods takes the form: `comObj$methodName(arg1,arg2,arg3,...)`
* Setting properties takes the form: `myObj[["Property"]] = TRUE`

With these basic capabilities, we can access and manipulate all the objects exposed in the PowerPoint API.

Let's create a new PowerPoint presentation from R  

```r
# Start up PowerPoint 
pp <- COMCreate("Powerpoint.Application")

# Make the application visible
pp[["Visible"]] = 1

# Add a new presentation
presentation <- pp[["Presentations"]]$Add()

# The presentation is empty.  Add a slide to it.
slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutBlank)
     
```

The critical part is, of course, knowing what methods and properties are available.  This is where the [PowerPoint 2013 Developer Reference][5] is handy, but not terribly user-friendly.  (You can also get this information from the Object Browser in the VBA editor.)

To understand how to go from VBA to R, let's recreate one of the basic examples in the documentation, [Applying Animations to Shapes in Office 2010][7], which works with Office 2013 as well.

The VBA code:

```vb.net
Sub TestPickupAnimation()
    With ActivePresentation.Slides(1)
        Dim shp1, shp2, shp3 As Shape
        ' Create the initial shape and apply the animations.
        Set shp1 = .Shapes.AddShape(msoShape12pointStar, 20, 20, 100, 100)
       
        .TimeLine.MainSequence.AddEffect shp1, msoAnimEffectFadedSwivel, , msoAnimTriggerAfterPrevious
        .TimeLine.MainSequence.AddEffect shp1, msoAnimEffectPathBounceRight, , msoAnimTriggerAfterPrevious
        .TimeLine.MainSequence.AddEffect shp1, msoAnimEffectSpin, , msoAnimTriggerAfterPrevious
       
        ' Now create a second shape, and apply the same animation to it:
        shp1.PickupAnimation
       
        Set shp2 = .Shapes.AddShape(msoShapeHexagon, 100, 20, 100, 100)
        shp2.ApplyAnimation
       
        ' And one more:
        Set shp3 = .Shapes.AddShape(msoShapeCloud, 180, 20, 100, 100)
        shp3.ApplyAnimation
       
    End With
End Sub

```


The code and the PowerPoint file created are available from [Github](https://github.com/asifsalam/PowerPoint_from_R).

Coming up:  
Part 2 - We'll scrape some data using [`XML`][3] and [`rvest`][4] and create a dataset of Clint Eastwood's movies.  
Part 3 - Then some fun! We'll use the data to play around with more advanced animation and interaction in PowerPoint.

[1]:"{{ site.url }} /R-and-PowerPoint-Part-1/"
[2]:http://www.omegahat.org/RDCOMClient/ "RCDOMClient"
[3]:http://www.omegahat.org/RSXML/ "XML Package for R"
[4]:https://github.com/hadley/rvest "rvest"
[5]:https://msdn.microsoft.com/en-us/library/office/ee861525.aspx "PowerPoint 2013 Developer Reference"
[6]:https://msdn.microsoft.com/en-us/library/office/ff743835.aspx "PowePoint Object Model Reference"
[7]:https://msdn.microsoft.com/en-us/library/office/gg190747(v=office.14).aspx "Applying Animations to Shapes"
[8]:https://msdn.microsoft.com/en-us/library/office/aa211626(v=office.11).aspx "AddEffect Method"
[9]:https://msdn.microsoft.com/en-us/library/office/ff744177.aspx "Shape Oject"
[10]:https://msdn.microsoft.com/EN-US/library/office/ff745204.aspx "TextFrame Property"
[11]:https://msdn.microsoft.com/EN-US/library/office/ff746145.aspx "Fill Property"
[12]:https://github.com/asifsalam/PowerPoint_from_R/blob/master/mso.txt "mso.txt"
[13]:http://www.imdb.org/ "IMDB"