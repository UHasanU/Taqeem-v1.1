from Taqeem import *

taqeem = Taqeem()

# Add the Excel filename containing the taqeemat followed by the worksheet name
taqeem.add_resource("PHYSICS.xlsx", "ردود النموذج 1")

# Set the column of the instructors in the worksheet
taqeem.set_instructors_column("C")

# Set the columns, in order, that you want to display in the table from the worksheet
taqeem.set_columns_to_display("B", "K", "D", "F", "G", "H", "I", "J")

# Add the title page
taqeem.add_title_page("تقييمات الفيزياء")

# Add the taqeemat
taqeem.add_taqeemat()

# Save the file with the name: PHYSICS
taqeem.save("PHYSICS.docx")


