namespace OutlookGUI
{
    partial class GUI
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.updateButton = new System.Windows.Forms.Button();
            this.recipientUpdateGroup = new System.Windows.Forms.GroupBox();
            this.recipientUpdateTextBox = new System.Windows.Forms.TextBox();
            this.updateDetailsGroup = new System.Windows.Forms.GroupBox();
            this.durationUpdateTextBox = new System.Windows.Forms.TextBox();
            this.durationUpdateLabel = new System.Windows.Forms.Label();
            this.startDateTextBox = new System.Windows.Forms.TextBox();
            this.dateUpdateLabel = new System.Windows.Forms.Label();
            this.subjectUpdateTextBox = new System.Windows.Forms.TextBox();
            this.updateSubjectLabel = new System.Windows.Forms.Label();
            this.eventDetailsGroup = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dateUpdateTextBox = new System.Windows.Forms.TextBox();
            this.eventDateLabel = new System.Windows.Forms.Label();
            this.subjectEventTextBox = new System.Windows.Forms.TextBox();
            this.subjectUpdateLabel = new System.Windows.Forms.Label();
            this.calendarUpdateGroup = new System.Windows.Forms.GroupBox();
            this.otherUpdateTextBox = new System.Windows.Forms.TextBox();
            this.otherUpdateRadio = new System.Windows.Forms.RadioButton();
            this.mainUpdateRadio = new System.Windows.Forms.RadioButton();
            this.addEventPage = new System.Windows.Forms.TabPage();
            this.sendButton = new System.Windows.Forms.Button();
            this.recipientGroup = new System.Windows.Forms.GroupBox();
            this.recipientTextBox = new System.Windows.Forms.TextBox();
            this.meetingGroup = new System.Windows.Forms.GroupBox();
            this.iterationTextBox = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.startTextBox = new System.Windows.Forms.TextBox();
            this.startLabel = new System.Windows.Forms.Label();
            this.recurrenceGroup = new System.Windows.Forms.GroupBox();
            this.endDateTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.yearlyRadio = new System.Windows.Forms.RadioButton();
            this.monthlyRadio = new System.Windows.Forms.RadioButton();
            this.weeklyRadio = new System.Windows.Forms.RadioButton();
            this.dailyRadio = new System.Windows.Forms.RadioButton();
            this.noneRadio = new System.Windows.Forms.RadioButton();
            this.durationTextBox = new System.Windows.Forms.TextBox();
            this.durationLabel = new System.Windows.Forms.Label();
            this.subjectTextBox = new System.Windows.Forms.TextBox();
            this.subjectLabel = new System.Windows.Forms.Label();
            this.calendarGroup = new System.Windows.Forms.GroupBox();
            this.otherTextBox = new System.Windows.Forms.TextBox();
            this.otherRadio = new System.Windows.Forms.RadioButton();
            this.mainRadio = new System.Windows.Forms.RadioButton();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.removeButton = new System.Windows.Forms.Button();
            this.eventDetailsGroupRemove = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.dateTextBoxRemove = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.subjectTextBoxRemove = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.calendarGroupRemove = new System.Windows.Forms.GroupBox();
            this.otherTextBoxRemove = new System.Windows.Forms.TextBox();
            this.otherRadioRemove = new System.Windows.Forms.RadioButton();
            this.mainRadioRemove = new System.Windows.Forms.RadioButton();
            this.tabPage2.SuspendLayout();
            this.recipientUpdateGroup.SuspendLayout();
            this.updateDetailsGroup.SuspendLayout();
            this.eventDetailsGroup.SuspendLayout();
            this.calendarUpdateGroup.SuspendLayout();
            this.addEventPage.SuspendLayout();
            this.recipientGroup.SuspendLayout();
            this.meetingGroup.SuspendLayout();
            this.recurrenceGroup.SuspendLayout();
            this.calendarGroup.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.eventDetailsGroupRemove.SuspendLayout();
            this.calendarGroupRemove.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.LightGray;
            this.tabPage2.Controls.Add(this.updateButton);
            this.tabPage2.Controls.Add(this.recipientUpdateGroup);
            this.tabPage2.Controls.Add(this.updateDetailsGroup);
            this.tabPage2.Controls.Add(this.eventDetailsGroup);
            this.tabPage2.Controls.Add(this.calendarUpdateGroup);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(513, 437);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Update Event";
            // 
            // updateButton
            // 
            this.updateButton.Location = new System.Drawing.Point(163, 317);
            this.updateButton.Name = "updateButton";
            this.updateButton.Size = new System.Drawing.Size(176, 60);
            this.updateButton.TabIndex = 4;
            this.updateButton.Text = "Update";
            this.updateButton.UseVisualStyleBackColor = true;
            this.updateButton.Click += new System.EventHandler(this.updateButton_Click);
            // 
            // recipientUpdateGroup
            // 
            this.recipientUpdateGroup.Controls.Add(this.recipientUpdateTextBox);
            this.recipientUpdateGroup.Location = new System.Drawing.Point(267, 9);
            this.recipientUpdateGroup.Name = "recipientUpdateGroup";
            this.recipientUpdateGroup.Size = new System.Drawing.Size(232, 285);
            this.recipientUpdateGroup.TabIndex = 3;
            this.recipientUpdateGroup.TabStop = false;
            this.recipientUpdateGroup.Text = "Recipients (Name or Email - 1 per line)";
            // 
            // recipientUpdateTextBox
            // 
            this.recipientUpdateTextBox.Location = new System.Drawing.Point(6, 19);
            this.recipientUpdateTextBox.Multiline = true;
            this.recipientUpdateTextBox.Name = "recipientUpdateTextBox";
            this.recipientUpdateTextBox.Size = new System.Drawing.Size(220, 258);
            this.recipientUpdateTextBox.TabIndex = 0;
            // 
            // updateDetailsGroup
            // 
            this.updateDetailsGroup.Controls.Add(this.durationUpdateTextBox);
            this.updateDetailsGroup.Controls.Add(this.durationUpdateLabel);
            this.updateDetailsGroup.Controls.Add(this.startDateTextBox);
            this.updateDetailsGroup.Controls.Add(this.dateUpdateLabel);
            this.updateDetailsGroup.Controls.Add(this.subjectUpdateTextBox);
            this.updateDetailsGroup.Controls.Add(this.updateSubjectLabel);
            this.updateDetailsGroup.Location = new System.Drawing.Point(8, 189);
            this.updateDetailsGroup.Name = "updateDetailsGroup";
            this.updateDetailsGroup.Size = new System.Drawing.Size(246, 105);
            this.updateDetailsGroup.TabIndex = 2;
            this.updateDetailsGroup.TabStop = false;
            this.updateDetailsGroup.Text = "Update Details";
            // 
            // durationUpdateTextBox
            // 
            this.durationUpdateTextBox.Location = new System.Drawing.Point(63, 77);
            this.durationUpdateTextBox.Name = "durationUpdateTextBox";
            this.durationUpdateTextBox.Size = new System.Drawing.Size(166, 20);
            this.durationUpdateTextBox.TabIndex = 8;
            // 
            // durationUpdateLabel
            // 
            this.durationUpdateLabel.AutoSize = true;
            this.durationUpdateLabel.Location = new System.Drawing.Point(7, 80);
            this.durationUpdateLabel.Name = "durationUpdateLabel";
            this.durationUpdateLabel.Size = new System.Drawing.Size(57, 15);
            this.durationUpdateLabel.TabIndex = 7;
            this.durationUpdateLabel.Text = "Duration:";
            // 
            // startDateTextBox
            // 
            this.startDateTextBox.Location = new System.Drawing.Point(63, 51);
            this.startDateTextBox.Name = "startDateTextBox";
            this.startDateTextBox.Size = new System.Drawing.Size(166, 20);
            this.startDateTextBox.TabIndex = 6;
            this.startDateTextBox.Text = "11/10/2017 6:00:00 PM";
            // 
            // dateUpdateLabel
            // 
            this.dateUpdateLabel.AutoSize = true;
            this.dateUpdateLabel.Location = new System.Drawing.Point(2, 54);
            this.dateUpdateLabel.Name = "dateUpdateLabel";
            this.dateUpdateLabel.Size = new System.Drawing.Size(64, 15);
            this.dateUpdateLabel.TabIndex = 5;
            this.dateUpdateLabel.Text = "Start Date:";
            // 
            // subjectUpdateTextBox
            // 
            this.subjectUpdateTextBox.Location = new System.Drawing.Point(63, 25);
            this.subjectUpdateTextBox.Name = "subjectUpdateTextBox";
            this.subjectUpdateTextBox.Size = new System.Drawing.Size(166, 20);
            this.subjectUpdateTextBox.TabIndex = 4;
            // 
            // updateSubjectLabel
            // 
            this.updateSubjectLabel.AutoSize = true;
            this.updateSubjectLabel.Location = new System.Drawing.Point(11, 28);
            this.updateSubjectLabel.Name = "updateSubjectLabel";
            this.updateSubjectLabel.Size = new System.Drawing.Size(51, 15);
            this.updateSubjectLabel.TabIndex = 1;
            this.updateSubjectLabel.Text = "Subject:";
            // 
            // eventDetailsGroup
            // 
            this.eventDetailsGroup.Controls.Add(this.label2);
            this.eventDetailsGroup.Controls.Add(this.dateUpdateTextBox);
            this.eventDetailsGroup.Controls.Add(this.eventDateLabel);
            this.eventDetailsGroup.Controls.Add(this.subjectEventTextBox);
            this.eventDetailsGroup.Controls.Add(this.subjectUpdateLabel);
            this.eventDetailsGroup.Location = new System.Drawing.Point(8, 89);
            this.eventDetailsGroup.Name = "eventDetailsGroup";
            this.eventDetailsGroup.Size = new System.Drawing.Size(247, 94);
            this.eventDetailsGroup.TabIndex = 1;
            this.eventDetailsGroup.TabStop = false;
            this.eventDetailsGroup.Text = "Event Details to Find";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(60, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(171, 15);
            this.label2.TabIndex = 6;
            this.label2.Text = "Leave empty for a single event";
            // 
            // dateUpdateTextBox
            // 
            this.dateUpdateTextBox.Location = new System.Drawing.Point(63, 49);
            this.dateUpdateTextBox.Name = "dateUpdateTextBox";
            this.dateUpdateTextBox.Size = new System.Drawing.Size(166, 20);
            this.dateUpdateTextBox.TabIndex = 5;
            this.dateUpdateTextBox.Text = "11/10/2017 6:00:00 PM";
            // 
            // eventDateLabel
            // 
            this.eventDateLabel.AutoSize = true;
            this.eventDateLabel.Location = new System.Drawing.Point(24, 52);
            this.eventDateLabel.Name = "eventDateLabel";
            this.eventDateLabel.Size = new System.Drawing.Size(36, 15);
            this.eventDateLabel.TabIndex = 4;
            this.eventDateLabel.Text = "Date:";
            // 
            // subjectEventTextBox
            // 
            this.subjectEventTextBox.Location = new System.Drawing.Point(63, 23);
            this.subjectEventTextBox.Name = "subjectEventTextBox";
            this.subjectEventTextBox.Size = new System.Drawing.Size(166, 20);
            this.subjectEventTextBox.TabIndex = 3;
            // 
            // subjectUpdateLabel
            // 
            this.subjectUpdateLabel.AutoSize = true;
            this.subjectUpdateLabel.Location = new System.Drawing.Point(11, 26);
            this.subjectUpdateLabel.Name = "subjectUpdateLabel";
            this.subjectUpdateLabel.Size = new System.Drawing.Size(51, 15);
            this.subjectUpdateLabel.TabIndex = 0;
            this.subjectUpdateLabel.Text = "Subject:";
            // 
            // calendarUpdateGroup
            // 
            this.calendarUpdateGroup.Controls.Add(this.otherUpdateTextBox);
            this.calendarUpdateGroup.Controls.Add(this.otherUpdateRadio);
            this.calendarUpdateGroup.Controls.Add(this.mainUpdateRadio);
            this.calendarUpdateGroup.Location = new System.Drawing.Point(8, 9);
            this.calendarUpdateGroup.Name = "calendarUpdateGroup";
            this.calendarUpdateGroup.Size = new System.Drawing.Size(247, 74);
            this.calendarUpdateGroup.TabIndex = 0;
            this.calendarUpdateGroup.TabStop = false;
            this.calendarUpdateGroup.Text = "Calendar Options";
            // 
            // otherUpdateTextBox
            // 
            this.otherUpdateTextBox.Enabled = false;
            this.otherUpdateTextBox.Location = new System.Drawing.Point(63, 42);
            this.otherUpdateTextBox.Name = "otherUpdateTextBox";
            this.otherUpdateTextBox.Size = new System.Drawing.Size(166, 20);
            this.otherUpdateTextBox.TabIndex = 2;
            // 
            // otherUpdateRadio
            // 
            this.otherUpdateRadio.AutoSize = true;
            this.otherUpdateRadio.Location = new System.Drawing.Point(6, 43);
            this.otherUpdateRadio.Name = "otherUpdateRadio";
            this.otherUpdateRadio.Size = new System.Drawing.Size(58, 19);
            this.otherUpdateRadio.TabIndex = 1;
            this.otherUpdateRadio.TabStop = true;
            this.otherUpdateRadio.Text = "Other";
            this.otherUpdateRadio.UseVisualStyleBackColor = true;
            this.otherUpdateRadio.CheckedChanged += new System.EventHandler(this.otherUpdateRadio_CheckedChanged);
            // 
            // mainUpdateRadio
            // 
            this.mainUpdateRadio.AutoSize = true;
            this.mainUpdateRadio.Checked = true;
            this.mainUpdateRadio.Location = new System.Drawing.Point(6, 20);
            this.mainUpdateRadio.Name = "mainUpdateRadio";
            this.mainUpdateRadio.Size = new System.Drawing.Size(109, 19);
            this.mainUpdateRadio.TabIndex = 0;
            this.mainUpdateRadio.TabStop = true;
            this.mainUpdateRadio.Text = "Main Calendar";
            this.mainUpdateRadio.UseVisualStyleBackColor = true;
            // 
            // addEventPage
            // 
            this.addEventPage.BackColor = System.Drawing.Color.LightGray;
            this.addEventPage.Controls.Add(this.sendButton);
            this.addEventPage.Controls.Add(this.recipientGroup);
            this.addEventPage.Controls.Add(this.meetingGroup);
            this.addEventPage.Controls.Add(this.calendarGroup);
            this.addEventPage.Location = new System.Drawing.Point(4, 22);
            this.addEventPage.Name = "addEventPage";
            this.addEventPage.Padding = new System.Windows.Forms.Padding(3);
            this.addEventPage.Size = new System.Drawing.Size(513, 437);
            this.addEventPage.TabIndex = 0;
            this.addEventPage.Text = "Add Event";
            // 
            // sendButton
            // 
            this.sendButton.Location = new System.Drawing.Point(327, 379);
            this.sendButton.Name = "sendButton";
            this.sendButton.Size = new System.Drawing.Size(145, 41);
            this.sendButton.TabIndex = 3;
            this.sendButton.Text = "Send";
            this.sendButton.UseVisualStyleBackColor = true;
            this.sendButton.Click += new System.EventHandler(this.sendButton_Click);
            // 
            // recipientGroup
            // 
            this.recipientGroup.Controls.Add(this.recipientTextBox);
            this.recipientGroup.Location = new System.Drawing.Point(282, 14);
            this.recipientGroup.Name = "recipientGroup";
            this.recipientGroup.Size = new System.Drawing.Size(221, 356);
            this.recipientGroup.TabIndex = 2;
            this.recipientGroup.TabStop = false;
            this.recipientGroup.Text = "Recipients (Name or Email) - 1 Per Line";
            // 
            // recipientTextBox
            // 
            this.recipientTextBox.Location = new System.Drawing.Point(9, 21);
            this.recipientTextBox.Multiline = true;
            this.recipientTextBox.Name = "recipientTextBox";
            this.recipientTextBox.Size = new System.Drawing.Size(200, 318);
            this.recipientTextBox.TabIndex = 0;
            // 
            // meetingGroup
            // 
            this.meetingGroup.Controls.Add(this.iterationTextBox);
            this.meetingGroup.Controls.Add(this.label6);
            this.meetingGroup.Controls.Add(this.startTextBox);
            this.meetingGroup.Controls.Add(this.startLabel);
            this.meetingGroup.Controls.Add(this.recurrenceGroup);
            this.meetingGroup.Controls.Add(this.durationTextBox);
            this.meetingGroup.Controls.Add(this.durationLabel);
            this.meetingGroup.Controls.Add(this.subjectTextBox);
            this.meetingGroup.Controls.Add(this.subjectLabel);
            this.meetingGroup.Location = new System.Drawing.Point(8, 85);
            this.meetingGroup.Name = "meetingGroup";
            this.meetingGroup.Size = new System.Drawing.Size(264, 336);
            this.meetingGroup.TabIndex = 1;
            this.meetingGroup.TabStop = false;
            this.meetingGroup.Text = "Meeting Details";
            // 
            // iterationTextBox
            // 
            this.iterationTextBox.Location = new System.Drawing.Point(63, 105);
            this.iterationTextBox.Name = "iterationTextBox";
            this.iterationTextBox.Size = new System.Drawing.Size(175, 20);
            this.iterationTextBox.TabIndex = 10;
            this.iterationTextBox.Text = "1";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 108);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(51, 15);
            this.label6.TabIndex = 9;
            this.label6.Text = "# to add";
            // 
            // startTextBox
            // 
            this.startTextBox.Location = new System.Drawing.Point(63, 55);
            this.startTextBox.Name = "startTextBox";
            this.startTextBox.Size = new System.Drawing.Size(175, 20);
            this.startTextBox.TabIndex = 8;
            this.startTextBox.Text = "11/10/2017 6:00:00 PM";
            // 
            // startLabel
            // 
            this.startLabel.AutoSize = true;
            this.startLabel.Location = new System.Drawing.Point(25, 58);
            this.startLabel.Name = "startLabel";
            this.startLabel.Size = new System.Drawing.Size(35, 15);
            this.startLabel.TabIndex = 7;
            this.startLabel.Text = "Start:";
            // 
            // recurrenceGroup
            // 
            this.recurrenceGroup.Controls.Add(this.endDateTextBox);
            this.recurrenceGroup.Controls.Add(this.label1);
            this.recurrenceGroup.Controls.Add(this.yearlyRadio);
            this.recurrenceGroup.Controls.Add(this.monthlyRadio);
            this.recurrenceGroup.Controls.Add(this.weeklyRadio);
            this.recurrenceGroup.Controls.Add(this.dailyRadio);
            this.recurrenceGroup.Controls.Add(this.noneRadio);
            this.recurrenceGroup.Location = new System.Drawing.Point(6, 128);
            this.recurrenceGroup.Name = "recurrenceGroup";
            this.recurrenceGroup.Size = new System.Drawing.Size(258, 218);
            this.recurrenceGroup.TabIndex = 6;
            this.recurrenceGroup.TabStop = false;
            this.recurrenceGroup.Text = "Recurrence Option";
            // 
            // endDateTextBox
            // 
            this.endDateTextBox.Enabled = false;
            this.endDateTextBox.Location = new System.Drawing.Point(57, 190);
            this.endDateTextBox.Name = "endDateTextBox";
            this.endDateTextBox.Size = new System.Drawing.Size(175, 20);
            this.endDateTextBox.TabIndex = 9;
            this.endDateTextBox.Text = "11/10/2017";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(-1, 193);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 15);
            this.label1.TabIndex = 9;
            this.label1.Text = "End Date:";
            // 
            // yearlyRadio
            // 
            this.yearlyRadio.AutoSize = true;
            this.yearlyRadio.Location = new System.Drawing.Point(6, 156);
            this.yearlyRadio.Name = "yearlyRadio";
            this.yearlyRadio.Size = new System.Drawing.Size(61, 19);
            this.yearlyRadio.TabIndex = 4;
            this.yearlyRadio.TabStop = true;
            this.yearlyRadio.Text = "Yearly";
            this.yearlyRadio.UseVisualStyleBackColor = true;
            this.yearlyRadio.CheckedChanged += new System.EventHandler(this.yearlyRadio_CheckedChanged);
            // 
            // monthlyRadio
            // 
            this.monthlyRadio.AutoSize = true;
            this.monthlyRadio.Location = new System.Drawing.Point(6, 124);
            this.monthlyRadio.Name = "monthlyRadio";
            this.monthlyRadio.Size = new System.Drawing.Size(71, 19);
            this.monthlyRadio.TabIndex = 3;
            this.monthlyRadio.TabStop = true;
            this.monthlyRadio.Text = "Monthly";
            this.monthlyRadio.UseVisualStyleBackColor = true;
            this.monthlyRadio.CheckedChanged += new System.EventHandler(this.monthlyRadio_CheckedChanged);
            // 
            // weeklyRadio
            // 
            this.weeklyRadio.AutoSize = true;
            this.weeklyRadio.Location = new System.Drawing.Point(6, 92);
            this.weeklyRadio.Name = "weeklyRadio";
            this.weeklyRadio.Size = new System.Drawing.Size(67, 19);
            this.weeklyRadio.TabIndex = 2;
            this.weeklyRadio.TabStop = true;
            this.weeklyRadio.Text = "Weekly";
            this.weeklyRadio.UseVisualStyleBackColor = true;
            this.weeklyRadio.CheckedChanged += new System.EventHandler(this.weeklyRadio_CheckedChanged);
            // 
            // dailyRadio
            // 
            this.dailyRadio.AutoSize = true;
            this.dailyRadio.Location = new System.Drawing.Point(6, 60);
            this.dailyRadio.Name = "dailyRadio";
            this.dailyRadio.Size = new System.Drawing.Size(55, 19);
            this.dailyRadio.TabIndex = 1;
            this.dailyRadio.TabStop = true;
            this.dailyRadio.Text = "Daily";
            this.dailyRadio.UseVisualStyleBackColor = true;
            this.dailyRadio.CheckedChanged += new System.EventHandler(this.dailyRadio_CheckedChanged);
            // 
            // noneRadio
            // 
            this.noneRadio.AutoSize = true;
            this.noneRadio.Checked = true;
            this.noneRadio.Location = new System.Drawing.Point(6, 28);
            this.noneRadio.Name = "noneRadio";
            this.noneRadio.Size = new System.Drawing.Size(58, 19);
            this.noneRadio.TabIndex = 0;
            this.noneRadio.TabStop = true;
            this.noneRadio.Text = "None";
            this.noneRadio.UseVisualStyleBackColor = true;
            this.noneRadio.CheckedChanged += new System.EventHandler(this.noneRadio_CheckedChanged);
            // 
            // durationTextBox
            // 
            this.durationTextBox.Location = new System.Drawing.Point(63, 79);
            this.durationTextBox.Name = "durationTextBox";
            this.durationTextBox.Size = new System.Drawing.Size(175, 20);
            this.durationTextBox.TabIndex = 5;
            // 
            // durationLabel
            // 
            this.durationLabel.AutoSize = true;
            this.durationLabel.Location = new System.Drawing.Point(10, 82);
            this.durationLabel.Name = "durationLabel";
            this.durationLabel.Size = new System.Drawing.Size(57, 15);
            this.durationLabel.TabIndex = 4;
            this.durationLabel.Text = "Duration:";
            // 
            // subjectTextBox
            // 
            this.subjectTextBox.Location = new System.Drawing.Point(63, 27);
            this.subjectTextBox.Name = "subjectTextBox";
            this.subjectTextBox.Size = new System.Drawing.Size(175, 20);
            this.subjectTextBox.TabIndex = 2;
            // 
            // subjectLabel
            // 
            this.subjectLabel.AutoSize = true;
            this.subjectLabel.Location = new System.Drawing.Point(14, 30);
            this.subjectLabel.Name = "subjectLabel";
            this.subjectLabel.Size = new System.Drawing.Size(51, 15);
            this.subjectLabel.TabIndex = 0;
            this.subjectLabel.Text = "Subject:";
            // 
            // calendarGroup
            // 
            this.calendarGroup.Controls.Add(this.otherTextBox);
            this.calendarGroup.Controls.Add(this.otherRadio);
            this.calendarGroup.Controls.Add(this.mainRadio);
            this.calendarGroup.Location = new System.Drawing.Point(8, 6);
            this.calendarGroup.Name = "calendarGroup";
            this.calendarGroup.Size = new System.Drawing.Size(264, 73);
            this.calendarGroup.TabIndex = 0;
            this.calendarGroup.TabStop = false;
            this.calendarGroup.Text = "Calendar Options";
            // 
            // otherTextBox
            // 
            this.otherTextBox.Enabled = false;
            this.otherTextBox.Location = new System.Drawing.Point(63, 41);
            this.otherTextBox.Name = "otherTextBox";
            this.otherTextBox.Size = new System.Drawing.Size(175, 20);
            this.otherTextBox.TabIndex = 2;
            // 
            // otherRadio
            // 
            this.otherRadio.AutoSize = true;
            this.otherRadio.Location = new System.Drawing.Point(6, 42);
            this.otherRadio.Name = "otherRadio";
            this.otherRadio.Size = new System.Drawing.Size(58, 19);
            this.otherRadio.TabIndex = 1;
            this.otherRadio.TabStop = true;
            this.otherRadio.Text = "Other";
            this.otherRadio.UseVisualStyleBackColor = true;
            this.otherRadio.CheckedChanged += new System.EventHandler(this.otherRadio_CheckedChanged);
            // 
            // mainRadio
            // 
            this.mainRadio.AutoSize = true;
            this.mainRadio.Checked = true;
            this.mainRadio.Location = new System.Drawing.Point(6, 19);
            this.mainRadio.Name = "mainRadio";
            this.mainRadio.Size = new System.Drawing.Size(109, 19);
            this.mainRadio.TabIndex = 0;
            this.mainRadio.TabStop = true;
            this.mainRadio.Text = "Main Calendar";
            this.mainRadio.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.addEventPage);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(0, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(521, 463);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.LightGray;
            this.tabPage1.Controls.Add(this.removeButton);
            this.tabPage1.Controls.Add(this.eventDetailsGroupRemove);
            this.tabPage1.Controls.Add(this.calendarGroupRemove);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(513, 437);
            this.tabPage1.TabIndex = 2;
            this.tabPage1.Text = "Remove Event";
            // 
            // removeButton
            // 
            this.removeButton.Location = new System.Drawing.Point(304, 65);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(175, 49);
            this.removeButton.TabIndex = 3;
            this.removeButton.Text = "Remove";
            this.removeButton.UseVisualStyleBackColor = true;
            this.removeButton.Click += new System.EventHandler(this.removeButton_Click);
            // 
            // eventDetailsGroupRemove
            // 
            this.eventDetailsGroupRemove.Controls.Add(this.label5);
            this.eventDetailsGroupRemove.Controls.Add(this.dateTextBoxRemove);
            this.eventDetailsGroupRemove.Controls.Add(this.label4);
            this.eventDetailsGroupRemove.Controls.Add(this.subjectTextBoxRemove);
            this.eventDetailsGroupRemove.Controls.Add(this.label3);
            this.eventDetailsGroupRemove.Location = new System.Drawing.Point(6, 85);
            this.eventDetailsGroupRemove.Name = "eventDetailsGroupRemove";
            this.eventDetailsGroupRemove.Size = new System.Drawing.Size(264, 93);
            this.eventDetailsGroupRemove.TabIndex = 2;
            this.eventDetailsGroupRemove.TabStop = false;
            this.eventDetailsGroupRemove.Text = "Event Details";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(42, 72);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(242, 15);
            this.label5.TabIndex = 6;
            this.label5.Text = "Leave empty to delete all instances of event";
            // 
            // dateTextBoxRemove
            // 
            this.dateTextBoxRemove.Location = new System.Drawing.Point(63, 49);
            this.dateTextBoxRemove.Name = "dateTextBoxRemove";
            this.dateTextBoxRemove.Size = new System.Drawing.Size(175, 20);
            this.dateTextBoxRemove.TabIndex = 5;
            this.dateTextBoxRemove.Text = "11/10/2017 6:00:00 PM";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(24, 52);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(36, 15);
            this.label4.TabIndex = 4;
            this.label4.Text = "Date:";
            // 
            // subjectTextBoxRemove
            // 
            this.subjectTextBoxRemove.Location = new System.Drawing.Point(63, 23);
            this.subjectTextBoxRemove.Name = "subjectTextBoxRemove";
            this.subjectTextBoxRemove.Size = new System.Drawing.Size(175, 20);
            this.subjectTextBoxRemove.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 15);
            this.label3.TabIndex = 0;
            this.label3.Text = "Subject:";
            // 
            // calendarGroupRemove
            // 
            this.calendarGroupRemove.Controls.Add(this.otherTextBoxRemove);
            this.calendarGroupRemove.Controls.Add(this.otherRadioRemove);
            this.calendarGroupRemove.Controls.Add(this.mainRadioRemove);
            this.calendarGroupRemove.Location = new System.Drawing.Point(6, 6);
            this.calendarGroupRemove.Name = "calendarGroupRemove";
            this.calendarGroupRemove.Size = new System.Drawing.Size(264, 73);
            this.calendarGroupRemove.TabIndex = 1;
            this.calendarGroupRemove.TabStop = false;
            this.calendarGroupRemove.Text = "Calendar Options";
            // 
            // otherTextBoxRemove
            // 
            this.otherTextBoxRemove.Enabled = false;
            this.otherTextBoxRemove.Location = new System.Drawing.Point(63, 41);
            this.otherTextBoxRemove.Name = "otherTextBoxRemove";
            this.otherTextBoxRemove.Size = new System.Drawing.Size(175, 20);
            this.otherTextBoxRemove.TabIndex = 2;
            // 
            // otherRadioRemove
            // 
            this.otherRadioRemove.AutoSize = true;
            this.otherRadioRemove.Location = new System.Drawing.Point(6, 42);
            this.otherRadioRemove.Name = "otherRadioRemove";
            this.otherRadioRemove.Size = new System.Drawing.Size(58, 19);
            this.otherRadioRemove.TabIndex = 1;
            this.otherRadioRemove.TabStop = true;
            this.otherRadioRemove.Text = "Other";
            this.otherRadioRemove.UseVisualStyleBackColor = true;
            this.otherRadioRemove.CheckedChanged += new System.EventHandler(this.otherRadioRemove_CheckedChanged);
            // 
            // mainRadioRemove
            // 
            this.mainRadioRemove.AutoSize = true;
            this.mainRadioRemove.Checked = true;
            this.mainRadioRemove.Location = new System.Drawing.Point(6, 19);
            this.mainRadioRemove.Name = "mainRadioRemove";
            this.mainRadioRemove.Size = new System.Drawing.Size(109, 19);
            this.mainRadioRemove.TabIndex = 0;
            this.mainRadioRemove.TabStop = true;
            this.mainRadioRemove.Text = "Main Calendar";
            this.mainRadioRemove.UseVisualStyleBackColor = true;
            // 
            // GUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(518, 462);
            this.Controls.Add(this.tabControl1);
            this.Name = "GUI";
            this.Text = "Outlook Calendar Event Sender";
            this.tabPage2.ResumeLayout(false);
            this.recipientUpdateGroup.ResumeLayout(false);
            this.recipientUpdateGroup.PerformLayout();
            this.updateDetailsGroup.ResumeLayout(false);
            this.updateDetailsGroup.PerformLayout();
            this.eventDetailsGroup.ResumeLayout(false);
            this.eventDetailsGroup.PerformLayout();
            this.calendarUpdateGroup.ResumeLayout(false);
            this.calendarUpdateGroup.PerformLayout();
            this.addEventPage.ResumeLayout(false);
            this.recipientGroup.ResumeLayout(false);
            this.recipientGroup.PerformLayout();
            this.meetingGroup.ResumeLayout(false);
            this.meetingGroup.PerformLayout();
            this.recurrenceGroup.ResumeLayout(false);
            this.recurrenceGroup.PerformLayout();
            this.calendarGroup.ResumeLayout(false);
            this.calendarGroup.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.eventDetailsGroupRemove.ResumeLayout(false);
            this.eventDetailsGroupRemove.PerformLayout();
            this.calendarGroupRemove.ResumeLayout(false);
            this.calendarGroupRemove.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button updateButton;
        private System.Windows.Forms.GroupBox recipientUpdateGroup;
        private System.Windows.Forms.TextBox recipientUpdateTextBox;
        private System.Windows.Forms.GroupBox updateDetailsGroup;
        private System.Windows.Forms.TextBox durationUpdateTextBox;
        private System.Windows.Forms.Label durationUpdateLabel;
        private System.Windows.Forms.TextBox startDateTextBox;
        private System.Windows.Forms.Label dateUpdateLabel;
        private System.Windows.Forms.TextBox subjectUpdateTextBox;
        private System.Windows.Forms.Label updateSubjectLabel;
        private System.Windows.Forms.GroupBox eventDetailsGroup;
        private System.Windows.Forms.TextBox dateUpdateTextBox;
        private System.Windows.Forms.Label eventDateLabel;
        private System.Windows.Forms.TextBox subjectEventTextBox;
        private System.Windows.Forms.Label subjectUpdateLabel;
        private System.Windows.Forms.GroupBox calendarUpdateGroup;
        private System.Windows.Forms.TextBox otherUpdateTextBox;
        private System.Windows.Forms.RadioButton otherUpdateRadio;
        private System.Windows.Forms.RadioButton mainUpdateRadio;
        private System.Windows.Forms.TabPage addEventPage;
        private System.Windows.Forms.Button sendButton;
        private System.Windows.Forms.GroupBox recipientGroup;
        private System.Windows.Forms.TextBox recipientTextBox;
        private System.Windows.Forms.GroupBox meetingGroup;
        private System.Windows.Forms.TextBox startTextBox;
        private System.Windows.Forms.Label startLabel;
        private System.Windows.Forms.GroupBox recurrenceGroup;
        private System.Windows.Forms.TextBox endDateTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton yearlyRadio;
        private System.Windows.Forms.RadioButton monthlyRadio;
        private System.Windows.Forms.RadioButton weeklyRadio;
        private System.Windows.Forms.RadioButton dailyRadio;
        private System.Windows.Forms.RadioButton noneRadio;
        private System.Windows.Forms.TextBox durationTextBox;
        private System.Windows.Forms.Label durationLabel;
        private System.Windows.Forms.TextBox subjectTextBox;
        private System.Windows.Forms.Label subjectLabel;
        private System.Windows.Forms.GroupBox calendarGroup;
        private System.Windows.Forms.TextBox otherTextBox;
        private System.Windows.Forms.RadioButton otherRadio;
        private System.Windows.Forms.RadioButton mainRadio;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox calendarGroupRemove;
        private System.Windows.Forms.TextBox otherTextBoxRemove;
        private System.Windows.Forms.RadioButton otherRadioRemove;
        private System.Windows.Forms.RadioButton mainRadioRemove;
        private System.Windows.Forms.GroupBox eventDetailsGroupRemove;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox dateTextBoxRemove;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox subjectTextBoxRemove;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button removeButton;
        private System.Windows.Forms.TextBox iterationTextBox;
        private System.Windows.Forms.Label label6;
    }
}

