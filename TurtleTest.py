import turtle

max_length = 500
increment = 10

def draw_spiral(a_turtle, line_length):
    if line_length > max_length:
        return
    a_turtle.forward(line_length)
    a_turtle.right(90)
    # recursive function call
    draw_spiral(a_turtle, line_length + increment)


tortuga = turtle.Turtle(shape="turtle")
tortuga.pensize(2)
tortuga.speed(10)
tortuga.color("red")
draw_spiral(tortuga, 10)
turtle.done()