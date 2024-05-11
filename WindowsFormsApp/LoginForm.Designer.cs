
namespace WindowsFormsApp
{
    partial class LoginForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_account = new System.Windows.Forms.TextBox();
            this.tb_password = new System.Windows.Forms.TextBox();
            this.btn_login = new System.Windows.Forms.Button();
            this.CB_Showpass = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "ERP Account:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "ERP Password:";
            // 
            // tb_account
            // 
            this.tb_account.Location = new System.Drawing.Point(93, 21);
            this.tb_account.Name = "tb_account";
            this.tb_account.Size = new System.Drawing.Size(145, 20);
            this.tb_account.TabIndex = 2;
            this.tb_account.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tb_account_KeyDown);
            // 
            // tb_password
            // 
            this.tb_password.Location = new System.Drawing.Point(93, 55);
            this.tb_password.Name = "tb_password";
            this.tb_password.Size = new System.Drawing.Size(145, 20);
            this.tb_password.TabIndex = 3;
            this.tb_password.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TB_password_KeyDown);
            // 
            // btn_login
            // 
            this.btn_login.Location = new System.Drawing.Point(93, 115);
            this.btn_login.Name = "btn_login";
            this.btn_login.Size = new System.Drawing.Size(115, 59);
            this.btn_login.TabIndex = 4;
            this.btn_login.Text = "Login";
            this.btn_login.UseVisualStyleBackColor = true;
            this.btn_login.Click += new System.EventHandler(this.btn_login_Click);
            // 
            // CB_Showpass
            // 
            this.CB_Showpass.AutoSize = true;
            this.CB_Showpass.Location = new System.Drawing.Point(92, 81);
            this.CB_Showpass.Name = "CB_Showpass";
            this.CB_Showpass.Size = new System.Drawing.Size(101, 17);
            this.CB_Showpass.TabIndex = 5;
            this.CB_Showpass.Text = "Show password";
            this.CB_Showpass.UseVisualStyleBackColor = true;
            this.CB_Showpass.CheckedChanged += new System.EventHandler(this.checkBoxShowPassword_CheckedChanged);
            // 
            // LoginForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(288, 199);
            this.Controls.Add(this.CB_Showpass);
            this.Controls.Add(this.btn_login);
            this.Controls.Add(this.tb_password);
            this.Controls.Add(this.tb_account);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "LoginForm";
            this.Text = "LoginForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_account;
        private System.Windows.Forms.TextBox tb_password;
        private System.Windows.Forms.Button btn_login;
        private System.Windows.Forms.CheckBox CB_Showpass;
    }
}