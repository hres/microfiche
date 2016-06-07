select All_Products.AccessNum , Manufacturers.* , All_Products.MFRCode
from All_Products
join Manufacturers
  on Manufacturers.ManuCode like concat(All_Products.MFRCode,'%') WHERE AccessNum = 75737; 
