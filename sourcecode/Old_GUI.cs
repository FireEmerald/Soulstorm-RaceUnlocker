/*
 * ----------------------------------------------------------------------------
 * "THE BEER-WARE LICENSE" :
 * <Belial2003@gmail.com> wrote this file. As long as you retain this notice you
 * can do whatever you want with this stuff. If we meet some day, and you think
 * this stuff is worth it, you can buy me a beer in return. If you think we will not 
 * meet some day, you can also send me some. n0|Belial2003
 * ----------------------------------------------------------------------------
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SSRaceUnlocker
{
    public partial class GUI : Form
    {
        public GUI()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            Codebitch CB = new Codebitch();
            CB.SetClassicKey(ClassicBox.Text.ToString());
            CB.SetWAKey(WABox.Text.ToString());
            CB.SetDCKey(DCBox.Text.ToString());
            CB.RegWriter();

            MessageBox.Show("Send me beer and feedback if it worked!", "Have fun!");


        }

    }
}