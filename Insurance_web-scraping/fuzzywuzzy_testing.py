from fuzzywuzzy import fuzz

s1 = ["525 D2 Hatchback 5dr Geartronic 6sp 1.6DT",
      "525 D2 Hatchback 5dr PowerShift 6sp 1.6DT",
      "525 D4 Hatchback 5dr Geartronic 8sp 2.0DTTI",
      "525 D4 Hatchback 5dr Spts Auto 8sp 2.0DT",
      "525 D4 Luxury Hatchback 5dr Geartronic 8sp 2.0DTTI",
      "525 T4 Hatchback 5dr PowerShift 6sp 2.0T",
      "525 T4 Hatchback 5dr Spts Auto 6sp 1.6T [IMP]",
      "525 T4 Hatchback 5dr Spts Auto 6sp 2.0T",
      "525 T4 Luxuary Hatchback 5dr PowerShift 6sp 2.0T",
      "525 T5 R-Design Hatchback 5dr Geartronic 8sp 2.0T",
      "525 T5 R-Design Hatchback 5dr Spts Auto 6sp 2.5T",
      "526 Cross Country D4 Hatchback 5dr Geartronic 8sp 2.0DTTI",
      "526 Cross Country D4 Luxury Hatchback 5dr Spts Auto 6sp 2.0DT",
      "526 Cross Country T5 Hatchback 5dr Geartronic 8sp AWD 2.0T",
      "526 Cross Country T5 Luxury Hatchback 5dr Spts Auto 6sp AWD 2.5T"]

db_example = "Volvo 2014 V40 Cross Country Geartronic AWD 526 Hatchback 8sp Spts Auto 2.0"

s2 = "Volvo V40 Hatchback 2015 \n526 Cross Country T5 Hatchback 5dr Geartronic 8sp AWD 2.0T"

for i in range(len(s1)):
    print(s1[i], end=": ")
    print(fuzz.token_sort_ratio(s1[i], db_example))

print("\n")

print(fuzz.token_set_ratio(db_example, s2))