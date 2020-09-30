﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Exceleration.Forms
{
    public partial class InputForm : Form
    {
        public string TextInput { get; set; }
        public bool OkPressed { get; set; }

        public InputForm(string title)
        {
            InitializeComponent();
            this.Text = title;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            OkPressed = true;
            TextInput = InputTB.Text;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            OkPressed = false;
            this.Close();
        }
    }
}
