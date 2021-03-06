import matplotlib.pyplot as pyplot
import unicodedata


def draw_diagram(points_x, points_y, title, x_title, y_title, output_dir,
                 clear_diagram=True, color='black', add_legend=False, legend_title=''):

    if clear_diagram:
        pyplot.clf()

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
    del_x = (max_x - min_x) // 10
    min_x -= del_x
    max_x += del_x

    min_y = points_y[0]
    max_y = points_y[0]

    for y in points_y:
        if y < min_y:
            min_y = y
        if y > max_y:
            max_y = y
    del_y = (max_y - min_y) // 10
    min_y -= del_y
    max_y += del_y

    if add_legend:
        pyplot.legend()

    pyplot.axis([min_x, max_x, min_y, max_y])
    pyplot.plot(points_x, points_y, color=color, linestyle='solid', linewidth=1, label=legend_title)
    pyplot.savefig(output_dir)


def draw_box_diagram(data, plot_title, y_title, output_dir):

    pyplot.clf()

    pyplot.title(plot_title)
    pyplot.ylabel(y_title)

    pyplot.boxplot(data)
    pyplot.savefig(output_dir)

#draw_diagram([10, 15, 20], [3, 10, 3], 'a', 'b', 'c', 'd.pdf')
#draw_box_diagram([0,1,2,3,4,5,6,7,8,9,10],'MyPlot','x','y', 'myPlot.pdf')
#pyplot.boxplot([0,1,2,3,4,5,6,7,8,9,10])
#pyplot.show()