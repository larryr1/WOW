# WOW Downloader
Project for my school that enables the automatic downloading, processing, and presentation of a slideshow via an email.

## Background?
My school distributes morning announcements via a PowerPoint that is shown each morning. We also have several displays around campus we wanted to show the announcements on automatically, but we couldn't think of an easy to do that since it's a new slideshow each morning and could be send via an email attachment or OneDrive link. This project solves that, by using the Microsoft Graph API to read the emails in a service account's inbox. It downloads either the OneDrive file or email attachment, saves it, and runs it with the system's PowerPoint executable.

## Why call it WOW?
The slideshow used to be called the WOW, or Word of the Week. That came way before I was a student.

## The PowerPointTransformer.exe
This excutable uses the PowerPoint COM APIs to set a fixed transition on each slide, and to make it loop automatically. I lost the source for this file. However, [this](https://github.com/larryr1/WOW/blob/main/PowerPointTransformer_Source.cs) is what ILSpy tells me is in it.
