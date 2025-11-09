from PIL import Image
img = Image.new("RGB", (640, 480), (255,255,255) )  # mode, size, color(default: 0)
# ... (do something to img) ...
img.show()  # invoke image viewer for debugging
img.save("out.jpg")  # save

