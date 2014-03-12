####DISCLAIMER
I am not a VBA wizard. I decided that it's important for me to maintain this one legacy repository to remember my roots.

This scripts in this repository were developed over the course of about 2 months during my time working as a data analysis consultant for the user acquisition team of a social gaming company in Jan-Mar 2014.

The scripts were designed to allow the team to input raw data as recieved from 30+ advertising partners and several in-house data utilities and to process said data with a few mouse-clicks. The resulting data enables side-by-side comparisons between partners for both top-level cost data as well as a deeper analysis of user performance from each source. The scripts performed exactly as designed and were made to enable easy extension over the many months they were used until a suitable database solution could be built.

Note that these scripts were developed on a machine runnign OSX. As a result, several of these scripts make use of a module that ports to OSX a `Dictionary` with similar functionality to that found in VBA as implemented on Windows machines. I'm sure there's a very good reason that this isn't native in OSX VBA... The additional modules may be found here:

- Dictionary came from [this page](http://sysmod.wordpress.com/2011/11/02/dictionary-class-in-vba-instead-of-scripting-dictionary/).
- And the KeyValuePair helper from [this one](http://sysmod.wordpress.com/2011/11/24/dictionary-vba-class-update/).