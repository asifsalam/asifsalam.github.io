---
layout: post
title: Create amazing PowerPoint slides using R - Putting it together
---
###Part 3 (3)

Time to finally create that pesky PowerPoint slide. [Part 1][1] was about using RDCOM to manipulate PowerPoint, and in [Part 2][2], we scraped the data needed to create the slide.  

Let's put it all together.

##Setup
We begin by loading the necessary libraries:

```r
library(RDCOMClient)
library(plyr)
library(dplyr)
library(stringr)
library(rvest)
```

##Data and Images
We need the data and images we scraped in [Part 2][2].  To create the slide we need:  
1 - Clint Eastwood's filmography data  
2 - Poster images for all the films  
3 - Box office returns  

The code assumes that the data files and image sub-directory are in the current R working directory.

If you followed parts [1][1] & [2][2], the data we will use, <a href="https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/eastwood_films.tsv">Clint Eastwood's filmography (tsv)</a>, and <a href="https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/eastwood_box_office.tsv">box office receipts (tsv)</a> should be available in the R environment, and the film poster images stored in a folder in the working directory.  If not, the data and the images are available on [github][3].  You can download the repo and and get all the files or run the following code to get data and the film posters (and download poster images for three films manually):

```r
# Download and read in the data files
download.file("https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/eastwood_films.tsv", 
              destfile = "films.tsv")
download.file("https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/eastwood_box_office.tsv", 
              destfile = "box_office.tsv")
download.file("https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/mso.txt", 
              destfile = "mso.txt")

films <- read.table("films.csv",header=TRUE, stringsAsFactors=FALSE)
box_office <- read.table("box_office.csv",header=TRUE, stringsAsFactors=FALSE)
source("mso.txt")
actor_name <- "Clint Eastwood"
img_dir <- "img"
```

The films data frame includes the title, year of release, name of the character played by Mr. Eastwood, the url of the IMDB film page, and the local path to the poster image file.  The box_office data frame has the box office receipts for most of the films.

If you haven't downloaded the poster images, you can run the following code to scrape the images from IMDB.  Alternatively, clone the <a href="https://github.com/asifsalam/PowerPoint_from_R">github repository</a>.


```r
# If you haven't downloaded the images, already,
# Loop through the films and download the poster image into the "img" subdirectory.
# If the poster is not found, flag the file name with 0.
if (!file.exists("img")) dir.create("img")
for (i in 1:nrow(films)) {
    img_node <- html(films$url[i]) %>% 
        			html_nodes(xpath='//td[@id="img_primary"]//img')
    if (length(img_node)==0) {
        films$img_file[i] <- "img/img00.jpg"
        cat(i," : img file NOT FOUND: ",films$img_file[i],"\n")
		
    } 
    else {
        img_link <- html_attr(img_node,"src")
        cat(i," :",films$img_file[i]," : ",img_link,"\n")
        download.file(img_link,films$img_file[i],method="internal",mode="wb")
    }
}

# Check which of the files were not found and download them manually
films$title[which(films$img_file=="img/img00.jpg")]

# These images don't exist.  Download appropriate images manually, and rename
# Films, Star in the Dust, The First Traveling Sales Lady, Dumbo Pilot
films[55,"img_file"] <- "img/img55.jpg"
films[54,"img_file"] <- "img/img54.jpg"
films[52,"img_file"] <- "img/img52.jpg"

```

## Create the PowerPoint Slide
We can now construct the slide.  We begin by creating a blank slide.

```r
# Create the PowerPoint slide
pp <- COMCreate("Powerpoint.Application")
presentation <- pp[["Presentations"]]$Add()
#slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutBlank)
slide1 <- presentation[["Slides"]]$Add(1,ms$ppLayoutTitleOnly)
slide_width <- presentation[["PageSetup"]]$SlideWidth()
slide_height <- presentation[["PageSetup"]]$SlideHeight()

# Let's also create the rgb funtion
pp_rgb <- function(r,g,b) {
    return( r + g*256 + b*256^2)
}
```

Now let's set the slide background and some properties.  The [slide object][4] has a many attributes, but we'll just use a few.

```r

img_file <- gsub("/","\\\\",paste(getwd(),"/","img/clint_background_1.png",sep=""))
bg_image <- slide1[["Shapes"]]$AddPicture(img_file,TRUE,FALSE,0,0,slide_width,slide_height)
bg_rect <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,0, 0,slide_width, slide_height)
bg_rect_fill <- bg_rect[["Fill"]]
bg_rect_fill[["ForeColor"]][["RGB"]] <- pp_rgb(102, 25, 13)
bg_rect_fill[["Transparency"]] <- 0.1
bg_rect_line <- bg_rect[["Line"]]
bg_rect_line[["BackColor"]][["RGB"]] <- pp_rgb(102,25,13)

slide_title <- slide1[["Shapes"]][["Title"]]
slide_title_color <- slide1[["ColorScheme"]]$Colors(ms$ppTitle)
slide_title_color[["RGB"]] <- pp_rgb(243,211,129)
slide_title$ZOrder(ms$msoBringToFront)

```

Let's set up the slide title: **Filmogarphy: Clint Eastwood**

```r
# AutoSize: https://msdn.microsoft.com/EN-US/library/office/ff745311(v=office.15).aspx
slide_title_frame <- slide_title[["TextFrame"]]
slide_title_frame[["AutoSize"]] <- ms$ppAutoSizeNone
slide_title_frame[["AutoSize"]] <- ms$ppAutoSizeShapeToFitText

slide_title[["Top"]] <- 0
slide_title[["Left"]] <- 0
title_text <- slide_title[["TextFrame"]][["TextRange"]]
title_text[["Text"]] <- paste("Filmography: ",actor_name,sep="")
title_font <- title_text[["Font"]]
title_font[["Color"]][["RGB"]] <- pp_rgb(243,211,129)
title_font[["Size"]] <- 36
title_font[["Name"]] <- "Calibri"

```

Let's add some gratiuitous elements - a line and a circle that shows Mr. Eastwood's total box office earnings (till 2012).

```r
diameter <- 100
# Add a line
#line1 <- slide1[["Shapes"]]$AddLine(0,diameter/2,slide_width-diameter+2,diameter/2)
line1 <- slide1[["Shapes"]]$AddLine(0,diameter/2,slide_width,diameter/2)
line1_attr <- line1[["Line"]]
line1_attr[["Weight"]] <- 1
line1_attr[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

# Add a circle
circle1 <- slide1[["Shapes"]]$AddShape(ms$msoShapeOval,slide_width-diameter,0,diameter,diameter)
circle1[["Top"]] <- 0
circle1[["Left"]] <- slide_width - diameter
circle1[["Width"]] <- diameter
circle1[["Height"]] <- diameter
circle1_color <- circle1[["Fill"]]
circle1_color[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)
total_earnings <- format(sum(as.numeric(box_office$adjusted_gross))/1000000000,digits=3)
circle_text <- circle1[["TextFrame"]][["TextRange"]]
circle_text[["Text"]] <- paste(total_earnings,"BUSD","")
circle_font <- circle_text[["Font"]]
circle_font[["Color"]] <- pp_rgb(102,25,13)
circle_font[["Name"]] <- "Calibri"
circle_font[["Size"]] <- 24

earnings_text <- slide1[["Shapes"]]$AddTextbox(ms$msoTextOrientationHorizontal,
                                               slide_width-4*diameter,diameter/2-25,diameter*3+5,20)
earnings_range <- earnings_text[["TextFrame"]][["TextRange"]]
earnings_range[["Text"]] <- "Total Box Office Earnings"
earnings_font <- earnings_range[["Font"]]

earnings_font[["Color"]] <- pp_rgb(243,211,129)
earnings_font[["Size"]] <- 20
# When you are returning an object, you need to create a variable, and then set the properties
# This doesn't work
# earnings_range[["ParagraphFormat"]][["Alignment"]] <- ms$ppAlignRight
earnings_para <- earnings_range[["ParagraphFormat"]]
earnings_para[["Alignment"]] <- ms$ppAlignRight
```

We need some parameters in order to place the poster images on the slide.  
There are 60 images, so 20 images in three rows should work.

```r
num_cols <- 20
num_rows <- ceiling(nrow(films)/num_cols)

# Based on the number of rows and columns, calculate the image height and width
image_width=floor(slide_width/num_cols)
image_height=floor(slide_height/(num_rows+3))
image_offset <- 3

# We'll store reference to the images in order to manipulate them later
images <- list()
image <- NULL

```

Let's see how the images look on the slide.

```r
for (i in 1:nrow(films)) {
    x = 0 + image_width * ((i-1) %% num_cols)
    y = image_height*image_offset + image_height * ((i-1) %/% num_cols)
	# Assuming that the images are in the working directory
    image_file <- gsub("/","\\\\",paste(getwd(),"/",films$img_file[i],sep=""))
    images[[as.character(i)]] <- slide1[["Shapes"]]$AddPicture(image_file,TRUE,FALSE,x+1,y+1,image_width-2,image_height-2)
    image <- images[[as.character(i)]]
    
    line <- image[["Line"]]
    line[["Style"]] <- ms$msoLineSingle
    line[["Weight"]] <- 2
    line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)
}
```

<img style="float:right;max-width:45%;" class="centre_image" src="/images/pp_slide02.png">We now have a slide with three rows of film poster images.  

But the slide is static.  Let's add some animation!  We'll include the ability sort the poster images by release date and alphabeticailly by film title.  
We'll also add tooltips (the film title, name of the character played by Clint Eastwood, and the release year).  Clicking on the image should take you the film's IMDB page.

Let's delete all the images, and start with the two sort buttons.  These are standard rectangle shapes that will trigger the sorting.

```r
button_alpha <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 350,150,150,40)
btn <- button_alpha[["TextFrame"]][["TextRange"]]
btn[["Text"]] <- "Title"
btn_font <- btn[["Font"]]
btn_font[["Size"]] <- 18
btn_font[["Color"]] <- pp_rgb(243,211,129)

btn_fill <- button_alpha[["Fill"]]
btn_fill[["Visible"]] <-0
btn_rgb <- btn_fill[["ForeColor"]][["RGB"]]

btn_line <- button_alpha[["Line"]]
btn_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)

button_date <- slide1[["Shapes"]]$AddShape(ms$msoShapeRectangle,slide_width - 350+180,150,150,40)
btn <- button_date[["TextFrame"]][["TextRange"]]
btn[["Text"]] <- "Release Year"

btn_font <- btn[["Font"]]
btn_font[["Size"]] <- 18
btn_font[["Color"]] <- pp_rgb(243,211,129)

btn_fill <- button_date[["Fill"]]
btn_fill[["Visible"]] <-0
btn_rgb <- btn_fill[["ForeColor"]][["RGB"]]

btn_line <- button_date[["Line"]]
btn_line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)
```

As far as I can tell, this is the way animation seems to work in Powerpoint:
1 - An animation effect (fade, swivel etc.) is applied to a specific object
2 - The animation is added to a "timeline", which specifies the sequence in which the effect will be executed
3 - The animation can include how it is triggered, the duration and some effect specifice behaviours (such as the path)
4 - An external trigger can also be specified

Since we want to sort in two ways, we need to add two timeline sequences

```r
seq_alpha = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()
seq_date = slide1[["TimeLine"]][["InteractiveSequences"]]$Add()

```

We can create a function that will apply some type of animation to the shape, in this case the poster image.  The goal is to move the image from one point to another.
This function takes a timeline (sequence), the image to be animated, the trigger for the animation, the path along which the image will move and the duration, and applies the parameters to the image.

```r
animate_image <- function(seq,image,trigger,path,duration=1.5) {
    effect <- seq$AddEffect(Shape=image,effectID=ms$msoAnimEffectPathDown,
                            trigger=ms$msoAnimTriggerOnShapeClick)
    ani <- effect[["Behaviors"]]$Add(ms$msoAnimTypeMotion)
    aniMotionEffect <- ani[["MotionEffect"]]
    aniMotionEffect[["Path"]] <- path
    effectTiming <- effect[["Timing"]]
    effectTiming[["TriggerType"]] <- ms$msoAnimTriggerWithPrevious
    effectTiming[["TriggerShape"]] <- trigger
    effectTiming[["Duration"]] <- duration
}

```

Now we can add add animation to the images.  When you click on the "Alphabetical" button, the images should sort by film title, and when you click on the "Release Year" button, the images should sort by release year.  The tricky bit is to get the paths right.  We also add the url to film's IMDB home page, used when the image is clicked, and a tooltip, which shows the film title, character name and release year on hover.

```r

for (i in 1:nrow(films)) {  
    # Populate the slide with the poster images
    x = 0 + image_width * ((i-1) %% num_cols)
    y = image_height*3 + image_height * ((i-1) %/% num_cols)
    img_file <- gsub("/","\\\\",paste(getwd(),"/",films$img_file[i],sep=""))
    images[[as.character(i)]] <- slide1[["Shapes"]]$AddPicture(img_file,TRUE,FALSE,x+1,y+1,image_width-1,image_height-1)
    image <- images[[as.character(i)]]
    
    # Some formatting
    line <- image[["Line"]]
    line[["Style"]] <- ms$msoLineSingle
    line[["Weight"]] <- 2
    line[["ForeColor"]][["RGB"]] <- pp_rgb(205,255,243)
	line[["ForeColor"]][["RGB"]] <- pp_rgb(243,211,129)
    
    # Add the url to the IMDB page, and tool tip - the film title, character name and the release year
    link <- image$ActionSettings(ms$ppMouseClick)[["Hyperlink"]]
    link[["Address"]] <- films$url[i]
    link[["ScreenTip"]] <- paste0(films$title[i],"\nCharacter: ",films$character_name[i],"\nRelease Year: ",films$year[i])
    
    # Animate on button click, so the posters re-sort based on title or release year
    index <- which(films$title[order(films$title)]==films$title[i]) - 1
    l1 <- format((0 + image_width * (index %% num_cols) - x)/slide_width,digits=3)
    l2 <- format((image_height*3 + image_height * (index %/% num_cols) - y)/slide_height,digits=3)
    path <- paste0("M0,0 L",l1,",",l2)
    animate_image(seq_alpha,image,button_alpha,path)
    path <- paste0("M",l1,",",l2," L0,0")
    animate_image(seq_date,image,button_date,path)
}

```

While we now have a decent animated slide, it feels a bit incomplete and unsatisfactory.  If you click the wrong sort button, the images animate from the start of the path back to their current positions.

Download the [R code for this part here][6], and the [resulting PowerPoint slide here][7].  

The final version fixes this problem with a bit of a hack, making only the relevant sort button visible.  This version which includes a sortable bar chart of earnings per movie, and some initial animation.  You can get the code [here][8] and the slide [here][9] (there are no macros, only animation).  Also, see what happens when you click on one of the bars in the bar chart!

[Short video of the final version][16] and [how it was created.][15]

Previous posts:<br>
[Part 1][1] - Create amazing PowerPoint slides using R - The basics<br>
[Part 2][2] - Create amazing PowerPoint slides using R - Getting the data

[1]:http://asifsalam.github.io/R-and-PowerPoint-Part-1/ "PowerPoint from R"
[2]:http://asifsalam.github.io/R-and-PowerPoint-Part-2/ "PowerPoint from R"
[3]:https://github.com/asifsalam/PowerPoint_from_R "PowerPoint from R"
[4]:https://msdn.microsoft.com/EN-US/library/office/ff744643.aspx "Slide Object"
[5]:https://msdn.microsoft.com/EN-US/library/office/ff744153(v=office.15).aspx "Property enumerations"
[6]:https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/PowerPoint_R_Part_3.R "PowerPoint_R_Part_3"
[7]:https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/PowerPoint_R_Part_3.pptx "PowerPoint_R_Part_3b.pptx"
[8]:https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/PowerPoint_R_Final.R "PowerPoint_R_Final"
[9]:https://raw.githubusercontent.com/asifsalam/PowerPoint_from_R/master/PowerPoint_R_Final.pptx "Point_R_Final"
[15]:https://youtu.be/kCxBTPDiN08 "How to create amazing PowerPoint slides - video"
[16]:https://youtu.be/XoAanIO8zbM "PowerPoint Slide created using R - video"
