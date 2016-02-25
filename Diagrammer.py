import matplotlib.pyplot as pyplot

def draw_diagram(points_x, points_y, title, x_title, y_title, output_dir):

    pyplot.title(title)
    pyplot.xlabel(x_title)
    pyplot.ylabel(y_title)

    min_x = points_x[0]
    max_x = points_x[0]

    for x in points_x:
        if x < min_x:
            min_x = x
        if x > max_x:
            max_x = x

    min_y = points_y[0]
    max_y = points_y[0]

    for y in points_y:
        if y < min_y:
            min_y = y
        if y > max_y:
            max_y = y

    pyplot.axis([min_x, max_x, min_y, max_y])
    pyplot.plot(points_x, points_y, color='black', linestyle='solid', linewidth=1)
    pyplot.savefig(output_dir)


#draw_diagram([10, 15, 20], [3, 10, 3], 'a', 'b', 'c', 'd.pdf')