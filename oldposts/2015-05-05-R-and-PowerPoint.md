---
layout: post
title: PowerPoint from R - Create Amazing Slides
---

For many of us, there's just no getting away from PowerPoint.  So if you use R (or Python) for data analysis, and PowerPoint for presenting/distributing the results, it's worthwhile learning how to include this last step into your workflow.  In this talk, [S Anand uses Python and **pywin32**][1] to create some jaw dropping effects in Powerpoint.    

[**RDCOMClient**][2] by Duncan Temple Lang allows you to do the same thing using R. It provides the ability to "access and control applications such as Excel, Word, Power Point, Web browsers etc."  Beyond efficiency and repeatability, programmatic access enables you to do things that just would not be possible with point and click.  

<!-- <div class="warning" style="height:50px;border:1px solid #d9d9d9; background-color:#f2eeee; text-align:left; vertical-align:middle; "> -->
<div class="warning" style="background-color:#f2eeee;"><p>
<img class="centre_image" src="/images/caution_finland_road_sign_189.svg" alt="Caution" style="width:50px; margin:2px 0 0 5px; align:left;" ><b style="margin: 0 0 0 10px"><i>Gratuitous animation ahead</i></b></p>
</div>

In this tutorial, we'll learn how to create some elements of the S Anand's talk, and a few other things, with a focus on interaction and animation.  We'll learn the basics of accessing methods and properties of PowerPoint VBA objects, then scrape data on Clint Eastwood's movies from [IMDB][4], and use it to create a slide.  In the end, we should have a cool, fun PowerPoint slide with Clint Eastwood's filmography.  

(Clearly PowerPoint has a powerful animation engine, and the [sophisticated object model][6] allows you to programmatically manipulate almost everything. The [Microsoft documentation][5] seems to be organized as a challenge, at least for beginners.)

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

Load the packages:

```r  
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

The critical part is, of course, knowing what methods and properties are available.  This is where the [PowerPoint 2013 Developer Reference][5] is handy, but not terribly user-friendly.  As a start, let's recreate one of the examples there, [Applying Animations to Shapes in Office 2010][6], which works with Office 2013 as well.

The VBA Code is:

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

As you can see, Microsoft has defined a set of constants for each element of the presentation, such as shape, animation, trigger, etc.  I remember I had to search for a while to find a consolidated list of all the enumerated constants.  I created a file and just read them into one variable, `ms`.

In the ```ÀddEffect``` method above, you'll notice an "empty" argument.  That tripped me up for a while, since it doesn't work with R.
The [AddEffect method][7] shows the arguments the method expect ```(expression.AddEffect(Shape, effectId, Level, trigger, Index))```.  Using an explicitly named argument (trigger) works.

```
source("mso.txt")
shp1 <- slide1[["Shapes"]]$AddShape(ms$msoShape12pointStar,20,20,100,100)
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectFadedSwivel,
                                                        trigger=ms$msoAnimTriggerAfterPrevious)
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectPathBounceRight,
                                                        trigger=ms$msoAnimTriggerAfterPrevious)
slide1[["TimeLine"]][["MainSequence"]]$AddEffect(shp1,ms$msoAnimEffectSpin,
                                                        trigger=ms$msoAnimTriggerAfterPrevious)

shp1$PickupAnimation()
shp2 <- slide1[["Shapes"]]$AddShape(ms$msoShapeHexagon,100,20,100,100)
shp2$ApplyAnimation()

shp3 <- slide1[["Shapes"]]$AddShape(ms$msoShapeCloud,180,20,100,200)
shp3$ApplyAnimation()

```

This should create a presentation with the three shapes.  If you go into animation mode, the shapes should appear, with the animations triggered.

That's neat! We now know how to create a new presentation, add shapes and some animation.  

Next, we'll set background and fill colours (soon).




[1]:https://www.youtube.com/watch?v=aKCXj1DyEhM "S Anand"
[2]:http://www.omegahat.org/RDCOMClient/ "RCDOMClient"
[3]:http://www.omegahat.org/RSXML/ "XML Package for R"
[4]:http://www.imdb.org/ "IMDB"
[5]:https://msdn.microsoft.com/en-us/library/office/ee861525.aspx "PowerPoint 2013 Developer Reference"
[6]:https://msdn.microsoft.com/en-us/library/office/ff743835.aspx "PowePoint Object Model Reference"
[6]:https://msdn.microsoft.com/en-us/library/office/gg190747(v=office.14).aspx
[7]:https://msdn.microsoft.com/en-us/library/office/aa211626(v=office.11).aspx