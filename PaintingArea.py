from PySide6.QtWidgets import QWidget
from PySide6.QtGui import QPainter, QPen, QBrush, QImage
from PySide6.QtCore import Qt, QSize, QPoint, QRect


class PaintingArea(QWidget):
    def __init__(self, parent):
        super().__init__()

        self._parent = parent

        self.setMinimumSize(self._parent.size().width(), self._parent.size().height())

        """ 
        Buffer image. It is necessary in order to, when resizing, put their image into the buffer before resizing, 
        and insert it after resizing the widget. In this case, the image size is preserved, that is, it is not scaled.
        """
        self.buffer_image = QImage(0, 0, QImage.Format.Format_RGB32)

        # Setting up the main canvas
        self.image = QImage(self.width(), self.height(), QImage.Format.Format_RGB32)
        self.image.fill(Qt.GlobalColor.white)

        # Image stack size for Undo/Redo
        self.image_stack_limit = 50
        self.image_stack = list()
        self.image_stack.append(self.image.copy())
        self.current_stack_position = 0

        # Setting Default Tools
        self.painting = False
        self.pen_size = 3
        self.pen_color = Qt.GlobalColor.black
        self.pen_style = Qt.PenStyle.SolidLine
        self.pen_cap = Qt.PenCapStyle.RoundCap
        self.pen_join = Qt.PenJoinStyle.RoundJoin

        self.last_point = QPoint()

    def resizeEvent(self, event):

        # Save current image to buffer
        self.buffer_image = self.image

        # Adjust the canvas to the new window size and clear the canvas to avoid distortion
        self.image = self.image.scaled(self._parent.size().width(), self._parent.size().height())
        self.image.fill(Qt.GlobalColor.white)

        # Transfer the image from the buffer to the canvas, to the starting coordinate
        painter = QPainter(self.image)
        painter.drawImage(QPoint(0, 0), self.buffer_image)

    def mousePressEvent(self, event):

        if event.button() == Qt.MouseButton.LeftButton:
            painter = QPainter(self.image)
            painter.setPen(QPen(self.pen_color, self.pen_size, self.pen_style, self.pen_cap, self.pen_join))
            painter.drawPoint(event.pos())
            self.painting = True
            self.last_point = event.pos()

        self.update()

    def mouseMoveEvent(self, event):

        if (event.buttons() == Qt.MouseButton.LeftButton) and self.painting:
            painter = QPainter(self.image)
            painter.setPen(QPen(self.pen_color, self.pen_size, self.pen_style, self.pen_cap, self.pen_join))
            painter.drawLine(self.last_point, event.pos())
            self.last_point = event.pos()

        self.update()

    def mouseReleaseEvent(self, event):

        if event.button() == Qt.MouseButton.LeftButton:
            self.painting = False

            # Replacing an incorrectly sized zero (clean) image
            if len(self.image_stack) >= 1:
                temp_zero_img = self.image.copy()
                temp_zero_img.fill(Qt.GlobalColor.white)
                self.image_stack[0] = temp_zero_img.copy()

            """
            Adding elements to an image stack.
            As long as the stack size is not exceeded, elements are simply added to the end.
            If the size is exceeded, then the elements are shifted to the beginning by 1,
            and a new element is inserted at the end of the list.
            """

            if (len(self.image_stack) < self.image_stack_limit and
                    not (self.current_stack_position < len(self.image_stack) - 1)):

                self.image_stack.append(self.image.copy())
                self.current_stack_position = len(self.image_stack) - 1

                self.update()

            elif self.current_stack_position < len(self.image_stack) - 1:
                """
                If the current position is less than the number of elements (when Redo is performed),
                then the stack is rebuilt.
                """
                for i in range(len(self.image_stack) - 1, self.current_stack_position, -1):
                    self.image_stack.pop(i)

                self.image_stack.append(self.image.copy())

                self.current_stack_position = len(self.image_stack) - 1

            else:
                # Shift elements in a list
                self.image_stack.pop(0)
                # Replacing the last element (which was previously the first) with a new element
                self.image_stack.append(self.image.copy())

                self.current_stack_position = len(self.image_stack) - 1

            self.update()

    def paintEvent(self, event):

        canvas_painter = QPainter(self)

        canvas_painter.drawImage(QPoint(0, 0), self.image)

    def undo(self):

        # If the current position is not at the very minimum
        if self.current_stack_position > 0:
            self.current_stack_position -= 1

            self.image = self.image_stack[self.current_stack_position].copy()

            self.update()

    def redo(self):
        # If the current position is not at the very maximum of the stack
        if self.current_stack_position < len(self.image_stack) - 1:
            self.current_stack_position += 1

            self.image = self.image_stack[self.current_stack_position].copy()

            self.update()

    def keyPressEvent(self, event):
        print(event.key())

    def clear(self):

        # Reset current stack position
        self.current_stack_position = 0

        # Clear canvas
        self.image.fill(Qt.GlobalColor.white)

        # Copy clear canvas
        canvas = self.image.copy()

        # Clear Undo-Redo stack
        self.image_stack.clear()

        # Add zero image
        self.image_stack.append(canvas.copy())

        self.update()
