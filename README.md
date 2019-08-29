This is a script for compressing pptx files

Yes, Powerpoint has a compress funtion, but it only crop the picutres and decrease the resolution. It's a good way if all picutres you insterted are jpegs.

But the reality is, we usually just take a snapshot of screen and paste them in to the slides, they will be stored as tiff or emf files, which are extremely huge. compressing them to png and jpg will reduce the file size a lot.

this python script requires IMAGE MAGIC to compress the pictures

requirements:

1. ImageMagick
2. python3

preparation:

1. edit conf.py, set IMAGE_MAGIC_DIR to the correct folder
2. set TMP_DIR to a writable folder
3. run `python main.py path_of_your_slides.pptx`
4. it will create a `path_of_your_slides_compressed.pptx` file in the same folder
