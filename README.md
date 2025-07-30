# Product Roadmapping Tool

This Application pulls data from a Excel file and imports that data to create a complex, linear, product roadmap. The design behind this project closely resembles the theory within the MIT textbook "Technology Roadmapping and development" by Olivier Weck. 

Link: https://ssrc.mit.edu/technology-roadmapping-and-development-by-olivier-l-de-weck/


### Software Outline

This program will parse an Excel sheet (template within the Repo.) and pull data to create different charts. Each sheet within the file is related to different graphs within the end product (Either an interactive chart or an Image). 

The "Technology" sheet contains the data for the "Technology" chart within the graphic. The Y-Axis for this chart section is Impact Rating, provided by the user. The X-axis for this chart (and every chart in this graphic) is a unified timeline. The "Capabillities" sheet is similar to the "Technology" sheet in the way it is formatted, and the Y-Axis is the same. 

The "Roadmap" sheet is by-far the most complicated. This sheet contains all the data for the main roadmap chart at the top of the graphic. The major Y-Axis is Risk Rating, provided by user, and the minor Y-axis for different elements of the charts differ based on the chart element. The top point of each mark on this chart represents the (Time, Risk Rating) relation. The line extended from the bottom of this point is a scaled representation of the expected budget for the activity. The "Glass Box" around every marker is the "Overall Risk" of the activity. The X length of the Glass box represents the time range of which an activity is expected to be completed. The Y length of the box is the Max. budget expected for the activity.

Each activity, Capabillity, and technology has it's own identifier (Ex. A01, C02, T03). The first letter represnents what dataset the ID is from, Activity is A, Technology is T, and so-on. The number is the chronological order in which each figure is expected to happen within it's dataset.

<img width="1750" height="943" alt="{C91B5BB1-5C91-40E3-969D-FB991847C1C6}" src="https://github.com/user-attachments/assets/a4f8a3d3-3985-4922-9ef8-75b9b03561ef" />


