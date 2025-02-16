# importing the module
import imdb
  
# creating instance of IMDb
ia = imdb.IMDb()
  
# movie name
name = "Titanic"
  
# searchning the movie
search = ia.search_movie(name)
  
# printing the result
print(search)