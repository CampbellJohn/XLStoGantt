from PIL import Image, ImageDraw, ImageFont

""" To Do
    Sort out number line.
"""

class Gantt:
    def __init__(self, data):
        self.data = data
        self.settings(self.data)
        self.gantt_create()

    def settings(self, data):
        self.today = 48
        self.w = 960
        self.mar = 10
        self.l_col = 80
        self.rowwidth = 22

        if data[0][10] == "Q1":
            self.days = 92
            self.Q = "Q1"
        else:    
            self.days = 94
            self.Q = "Q2"

        self.canv = self.w - (self.mar * 2)
        self.r_col = self.canv - self.l_col
        self.dayp = int(self.r_col/self.days)
        self.topmar = self.dayp * 6

        if self.days % 2 == 0:
            self.halfdays = self.days/2
        else:
            self.halfdays = (self.days/2) + 1

        """ Set up Fonts """
        self.font = ImageFont.truetype("Ref\Regular.ttf", 14)
        self.font1 = ImageFont.truetype("Ref\Regular.ttf", 12)
        self.font2 = ImageFont.truetype("Ref\Regular.ttf", 10)
        self.font3 = ImageFont.truetype("Ref\Regular.ttf", 8)

        """ Set up Columns """
        self.n = range(self.canv)
        self.c = self.n[(self.l_col - self.dayp): self.canv: self.dayp]
        self.col = []
        self.pat = self.c[1::2]
        self.pattern = []

        for x in self.pat:
            self.pattern.append(x + self.mar)
        for x in self.c:
            self.col.append(x + self.mar)

        """ Set up Rows """
        self.vn = range(9999)
        self.vc = self.vn[(self.topmar): 9999: self.rowwidth]
        self.row = []
        for x in self.vc:
            self.row.append(x)

    def gantt_create(self):
        # Figure out how tall the image is, and add it to settings.
        self.heightcalc = int(len(self.data) * self.rowwidth) + \
        self.topmar + 20
        
        # Creates the image and draws the object.
        self.im = Image.new("RGBA", (self.w,self.heightcalc), color="#FFFFFF")

        # Draw the timeline.
        self.draw_timeline()

        if self.Q == "Q1":
            self.draw_month(1, 32, "January")
            self.draw_month(32, 61, "February")
            self.draw_month(61, 93, "March")
        if self.Q == "Q2":
            self.draw_month(1, 32, "April")
            self.draw_month(32, 64, "May")
            self.draw_month(64, 95, "June")
        
        # Sends rows to the bar function
        for x in range(len(self.data)):
            self.ganttbar( self.col[int(self.data[x][5] - self.data[x][6])], # X
                           self.row[x],          # Y
                           self.data[x][0],      # Text, 
                           int(self.data[x][6]), # dur
                           self.data[x][4],      # Tag,
                           self.data[x][5],      # Due Date Int
                           self.data[x][8]       # Completed?
                        )

        self.filename = "Gantt_" + str(self.data[x][1]) + "_" \
                        + str(self.data[x][10]) + "_" + str(self.data[x][9]) \
                        + ".png"
        
        # Write to stdout
        self.im.save("Final\\" + str(self.filename), "PNG")

    def draw_timeline(self):
        draw = ImageDraw.Draw(self.im)

        # Draw vertical bars
        for x in self.pattern[0: int(self.halfdays)]:
            draw.rectangle((x, self.topmar, (x + self.dayp), (self.heightcalc \
            -  10)), fill = "#C1E8E5")

        # Draw the "today" line
        if self.Q == "Q1":
            draw.rectangle((self.col[self.today], self.topmar, \
            (self.col[self.today] + self.dayp), (self.heightcalc \
            - 10)), fill = "#8A0917")

        # Outline the timeline
        draw.line(((self.mar + self.l_col), self.topmar, (self.mar \
        + self.l_col + (self.days * self.dayp)), self.topmar), fill = "#B8AE9C")
        draw.line(((self.mar + self.l_col), self.topmar, (self.mar \
        + self.l_col), (self.heightcalc - 10)), fill = "#B8AE9C")
        draw.line(((self.mar + self.l_col), self.topmar, (self.mar \
        + self.l_col), (self.heightcalc - 10)), fill = "#B8AE9C")
        draw.line(((self.mar + self.l_col + (self.days * self.dayp)), \
        self.topmar, (self.mar + self.l_col + (self.days*self.dayp)), \
        (self.heightcalc-10)), fill = "#B8AE9C")

        del draw

    """ Month Drawing Function """
    def draw_month(self, pos1, pos2, month):
        draw = ImageDraw.Draw(self.im)

        # Drawing the Month
        text_size = draw.textsize(month, font = self.font)
        barsize = self.col[pos2] - self.col[pos1]
        center = ((barsize - text_size[1]) / 2) + self.col[pos1]
        draw.text((center, (self.topmar-36)), month, font=self.font, \
        fill = "#B8AE9C")

        # Drawing numbers
        num = 1
        for x in self.c[pos1:pos2:7]:
            draw.text(((x + self.mar), (self.topmar - (self.rowwidth / 1.5))), \
            str(num), font = self.font2, fill = "#595241")

            if num == 1:
                num += 6
            else:
                num += 7

        # Horizontal Month Lines
        draw.line((self.col[pos1], (self.topmar - 18), self.col[pos2], \
        (self.topmar - 18)), fill = "#B8AE9C")
        draw.line((self.col[pos1], (self.topmar - 36), self.col[pos2], \
        (self.topmar - 36)), fill = "#B8AE9C")
        # Vertical Month Lines
        draw.line((self.col[pos1], (self.topmar - 36), self.col[pos1], \
        (self.topmar - 18)), fill = "#B8AE9C")
        draw.line((self.col[pos2], (self.topmar - 36), self.col[pos2], \
        (self.topmar - 18)), fill = "#B8AE9C")

        del draw

    """ Horizontal Bar Function """
    def ganttbar(self, x, y, text, dur, tag, due, comp):
        draw = ImageDraw.Draw(self.im)

        if self.Q == "Q2":
            due += 92

        # Determine bar color
        color = "#B8AE9C"
        if due < self.today:
            color = "#8A0917"
        if comp == "Y":
            color = "#595241"

        # Draw the bar
        draw.rectangle((x, y, (x + (self.dayp * dur)), (y + self.rowwidth)), \
        fill = color)

        # Font sizing
        textsize = draw.textsize(text, font = self.font)
        textsize1 = draw.textsize(text, font = self.font1)
        textsize2 = draw.textsize(text, font = self.font2)
        textsize3 = draw.textsize(text, font = self.font3)

        # Write text
        if textsize[0] < ((self.dayp * dur) - 10):
            draw.text(((x + 5), (y)), text, font = self.font, fill = "#FFFFFF")
        elif textsize1[0] < ((self.dayp * dur) - 5):
            draw.text(((x + 5), (y)), text, font = self.font1, fill = "#FFFFFF")
        elif textsize2[0] < ((self.dayp * dur) - 5):
            draw.text(((x + 5), (y)), text, font = self.font2, fill = "#FFFFFF")
        else:
            draw.text(((x + 5), (y)), text, font = self.font3, fill = "#FFFFFF")

        # Tag sizing
        tagsize = draw.textsize(tag, font = self.font)
        tagsize1 = draw.textsize(tag, font = self.font1)
        tagsize2 = draw.textsize(tag, font = self.font2)
        tagsize3 = draw.textsize(tag, font = self.font3)

        # Write tag
        tagx = self.col[0] - tagsize2[0]
        draw.text((tagx, y), tag, font = self.font2, fill = "#595241")

        del draw

"""
Campaign Gantt Object
    Have a row at the top for Webinars.
    A row/section/color for each product.

"""

class CampaignGantt(Gantt):
    def __init__(self, data):
        self.data = data
        self.settings(self.data)
        self.gantt_create()

    def settings(self, data):
        self.today = 48
        self.w = 960
        self.mar = 10
        self.l_col = 80
        self.rowwidth = 22

        if data[0][10] == "Q1":
            self.days = 92
            self.Q = "Q1"
        else:    
            self.days = 94
            self.Q = "Q2"

        self.canv = self.w - (self.mar * 2)
        self.r_col = self.canv - self.l_col
        self.dayp = int(self.r_col/self.days)
        self.topmar = self.dayp * 6

        if self.days % 2 == 0:
            self.halfdays = self.days/2
        else:
            self.halfdays = (self.days/2) + 1

        """ Set up Fonts """
        self.font = ImageFont.truetype("Ref\Regular.ttf", 14)
        self.font1 = ImageFont.truetype("Ref\Regular.ttf", 12)
        self.font2 = ImageFont.truetype("Ref\Regular.ttf", 10)
        self.font3 = ImageFont.truetype("Ref\Regular.ttf", 8)

        """ Set up Columns """
        self.n = range(self.canv)
        self.c = self.n[(self.l_col - self.dayp): self.canv: self.dayp]
        self.col = []
        self.pat = self.c[1::2]
        self.pattern = []

        for x in self.pat:
            self.pattern.append(x + self.mar)
        for x in self.c:
            self.col.append(x + self.mar)

        """ Set up Rows """
        self.vn = range(9999)
        self.vc = self.vn[(self.topmar): 9999: self.rowwidth]
        self.row = []
        for x in self.vc:
            self.row.append(x)

    def gantt_create(self):
        # Figure out how tall the image is, and add it to settings.
        self.heightcalc = int(len(self.data) * self.rowwidth) + \
        self.topmar + 20
        
        # Creates the image and draws the object.
        self.im = Image.new("RGBA", (self.w,self.heightcalc), color="#FFFFFF")

        # Draw the timeline.
        self.draw_timeline()

        if self.Q == "Q1":
            self.draw_month(1, 32, "January")
            self.draw_month(32, 61, "February")
            self.draw_month(61, 93, "March")
        if self.Q == "Q2":
            self.draw_month(1, 32, "April")
            self.draw_month(32, 64, "May")
            self.draw_month(64, 95, "June")
        
        # Sends rows to the bar function
        for x in range(len(self.data)):
            self.ganttbar( self.col[int(self.data[x][5] - self.data[x][6])], # X
                           self.row[x],          # Y
                           self.data[x][0],      # Text, 
                           int(self.data[x][6]), # dur
                           self.data[x][4],      # Tag,
                           self.data[x][5],      # Due Date Int
                           self.data[x][8]       # Completed?
                        )

        self.filename = "Gantt_" + str(self.data[x][1]) + "_" \
                        + str(self.data[x][10]) + "_" + str(self.data[x][9]) \
                        + ".png"
        
        # Write to stdout
        self.im.save("Final\\" + str(self.filename), "PNG")

