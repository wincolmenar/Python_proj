def BMI_Calculator():
    height = float(input("Please input your height in meters: "))
    weight= float(input("Please input your weight kilograms: "))

    bmi = weight / (height ** 2)

    result = ""

    if bmi < 18.5:
        result = "Underweight"
    elif bmi < 25:
        result = "Normal"
    elif bmi < 30:
        result = "Overweight"
    else:
        result = "Obese"

    print(f"Your BMI is {str(round(bmi, 2))} and you're {result}.")


BMI_Calculator()
