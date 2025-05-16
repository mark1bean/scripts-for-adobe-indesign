# Numbered Markers Quick Tutorial

Follow along this little story to see what happens.

---

*This is the minimal document set up:*

![Numbered Markers tutorial](images/numbered-markers-1.png)

---

![Numbered Markers script UI](images/numbered-markers-2.png)

---

![Numbered Markers script UI](images/numbered-markers-3.png)

Note: if you already have your document populated with "unmanaged" markers, and you don't want to reposition them, see [But I already have my markers all set up](#but-i-already-have-my-markers-all-set-up).

---

![Numbered Markers script UI](images/numbered-markers-4.png)

---

![Numbered Markers script UI](images/numbered-markers-5.png)

---

![Numbered Markers script UI](images/numbered-markers-6.png)

---

![Numbered Markers script UI](images/numbered-markers-7.png)

---

![Numbered Markers script UI](images/numbered-markers-8.png)

---

![Numbered Markers script UI](images/numbered-markers-9.png)

If there are more texts found than markers found, script will create markers for the "unmarkered" texts.

---

![Numbered Markers script UI](images/numbered-markers-10.png)

---

![Numbered Markers script UI](images/numbered-markers-11.png)

---

![Numbered Markers script UI](images/numbered-markers-12.png)

---

![Numbered Markers script UI](images/numbered-markers-13.png)

---

![Numbered Markers script UI](images/numbered-markers-14.png)

---

![Numbered Markers script UI](images/numbered-markers-15.png)

---

![Numbered Markers script UI](images/numbered-markers-16.png)

---

![Numbered Markers script UI](images/numbered-markers-17.png)

---

![Numbered Markers script UI](images/numbered-markers-18.png)

---

![Numbered Markers script UI](images/numbered-markers-19.png)

Here you can see how the script links between the numbered text and the marker. You can show the **Script Panel** by choosing menu *Window > Utilities > Script Label*.

---

![Numbered Markers script UI](images/numbered-markers-20.png)

---

![Numbered Markers script UI](images/numbered-markers-21.png)

---

## But I already have my markers all set up

The script makes a link between a numbered paragraph and its marker by matching the text of the paragraph to the script label of the marker. It does this automatically whenever it creates a marker, but there's no reason you can't do it manually.

There are two approaches here:

##### (a) Create new markers and reposition them

1. When you first run the script, turn OFF *Remove orphan markers*, but let it create a complete set of new markers.
1. Position the new markers *over the top* of the existing markers (don't bother deleting the old ones, it will happen in the next step).
1. Run the script again, this time with *Remove orphan markers* turned ON. This will remove all the non-linked markers.

##### (b) Link your markers to the texts, using Script Labels

1. Copy the exact text of the list item, eg "Silver Creek".
1. Select the correct marker for "Silver Creek".
1. In the Script Label panel (*Windows > Utilities > Script Label*), paste the text from step 1.
1. Repeat this for each marker.

Once you've done them all, you shouldn't have to do that process again—the script will handle it from then on.

---

## Download Script

[![Download Copy Things script](https://img.shields.io/badge/*_Download_Script_*-FREE!_-F50?style=flat-square)](https://raw.githubusercontent.com/mark1bean/scripts-for-adobe-indesign/main/Numbered%20Markers.js)   ![ExtendScript](https://img.shields.io/badge/Language-ExtendScript-99B?style=flat-square)   ![Version: 2025-05-15](https://img.shields.io/badge/Version-2025--05--15-5A5?style=flat-square)   [![Donate](https://img.shields.io/badge/Donate-PayPal-blue?style=flat-square)](https://www.paypal.com/donate?hosted_button_id=SBQHVWHSSTA9Q)

---

[Go Back](../README.md)
