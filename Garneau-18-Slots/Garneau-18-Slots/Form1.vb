Public Class frmPokeslots

    ' Created by: Andrew Garneau
    ' Program Name: 18-Slots
    ' Date Created: May 19th 2016
    ' Program Goal: A pokemon themed slot machine to play.

    ' sets all the global variables in the program. 
    ' creates a random variable
    Dim Rnd As New Random
    Dim image1 As Integer
    Dim image2 As Integer
    Dim image3 As Integer
    Dim cashin As String
    Dim credits As Double
    Dim inputcred As String
    Dim Bet As Double
    Dim Winnings As Double
    Dim Pulls As Double = 0
    Dim progressive As Double = 0
    ' subroutine that generates the 1st image.
    Sub GenImage1()

        ' sets the image 1 variable to a random number between 1 and 9 (1-8)
        image1 = Rnd.Next(1, 9)

        ' select case to display a certain image based on the value of image1
        Select Case image1
            ' if case = 1 then it shows a specific image. Repeats for all other images.
            Case 1
                picSlot1.Image = My.Resources.eevee
            Case 2
                picSlot1.Image = My.Resources.umbreon
            Case 3
                picSlot1.Image = My.Resources.espeon
            Case 4
                picSlot1.Image = My.Resources.jolteon
            Case 5
                picSlot1.Image = My.Resources.glaceon
            Case 6
                picSlot1.Image = My.Resources.flareon
            Case 7
                picSlot1.Image = My.Resources.leafeon
            Case 8
                picSlot1.Image = My.Resources.vaporeon
        End Select

    End Sub
    ' subroutine that generates the 2nd image.
    Sub GenImage2()

        ' sets the image 2 variable to a random number between 1 and 9 (1-8)
        image2 = Rnd.Next(1, 9)

        ' select case to display a certain image based on the value of image2
        Select Case image2
            ' if case = 1 then it shows a specific image. Repeats for all other images.
            Case 1
                picSlot2.Image = My.Resources.eevee
            Case 2
                picSlot2.Image = My.Resources.umbreon
            Case 3
                picSlot2.Image = My.Resources.espeon
            Case 4
                picSlot2.Image = My.Resources.jolteon
            Case 5
                picSlot2.Image = My.Resources.glaceon
            Case 6
                picSlot2.Image = My.Resources.flareon
            Case 7
                picSlot2.Image = My.Resources.leafeon
            Case 8
                picSlot2.Image = My.Resources.vaporeon
        End Select

    End Sub
    ' subroutine that generates the 3rd image.
    Sub GenImage3()

        ' sets the image 3 variable to a random number between 1 and 9 (1-8)
        image3 = Rnd.Next(1, 9)

        ' select case to display a certain image based on the value of image3
        Select Case image3
            ' if case = 1 then it shows a specific image. Repeats for all other images.
            Case 1
                picSlot3.Image = My.Resources.eevee
            Case 2
                picSlot3.Image = My.Resources.umbreon
            Case 3
                picSlot3.Image = My.Resources.espeon
            Case 4
                picSlot3.Image = My.Resources.jolteon
            Case 5
                picSlot3.Image = My.Resources.glaceon
            Case 6
                picSlot3.Image = My.Resources.flareon
            Case 7
                picSlot3.Image = My.Resources.leafeon
            Case 8
                picSlot3.Image = My.Resources.vaporeon
        End Select

    End Sub

    Private Sub btnPull_Click(sender As Object, e As EventArgs) Handles btnPull.Click

        ' sets a local variable to box buttons.
        ' used for message box when runs out of credits.
        Dim BoxButtons As Integer

        ' adds 1 to the pull variable each time the slot is pulled.
        Pulls += 1
        ' displays the amount of pulls.
        lblValPulls.Text = Pulls

        ' runs subroutine to generate images 1,2 and 3.
        GenImage1()
        GenImage2()
        GenImage3()

        ' subtracts the bet amount for credits each time pull is clicked.
        credits = credits - Bet
        ' displays credit amount after each pull (minus the bets)
        lblValTotalCred.Text = credits

        ' if statement to find out what happens after user runs out of credits.
        If credits < 0 Or credits = 0 Then

            'shows messagebox to ask if user wants to keep playing.
            BoxButtons = MessageBox.Show("Would you like to keep playing?", "Ran Out of Money!", MessageBoxButtons.YesNo, MessageBoxIcon.Error)

            ' if user clicks yes, runs loop to input more credits.
            If BoxButtons = DialogResult.Yes Then
                Do
                    ' displays input box for the cashin variable
                    cashin = InputBox("Please enter the amount of money you would like to put into the machine.", "Please input your credits")
                    ' if the cashin value isn't numeric, runs an error message.
                    If Not IsNumeric(cashin) Then
                        MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    End If
                    ' if the cashin variable is higher than 100,000, runs an error message.
                    If Val(cashin) > 100001 Then
                        MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    End If
                    ' if the cashin variable is 0 or less, runs an error message.
                    If Val(cashin) < 0 Then
                        MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                    End If
                    'loops until the cashin is a numeric value between 1 and 100000
                Loop Until Val(cashin) > 0 And Val(cashin) < 100001 And IsNumeric(cashin)

                ' muliplies cashin * 4 to convert it to credits, then addds it to the credit count.
                credits = credits + (cashin * 4)

            End If

            ' if no is clicked, then it displays a thank you message and total credit count, and ends the program.
            If BoxButtons = DialogResult.No Then

                ' displays the message box with their winnings (if any)
                MessageBox.Show("Thank you for playing! You won " & credits / 4 & " credits.", "Thank you!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

                ' ends the program
                End
            End If
        End If

        ' the upcoming code deals with winning matches.

        ' if all the images are 1 (same picture) then...
        If image1 = 1 And image2 = 1 And image3 = 1 Then
            ' winnings variable is set to the value of the match (this case the progressive) x the bet amount.
            Winnings = progressive * Bet
            ' displays the winning messagebox 
            MessageBox.Show("Congrats, you have won " & Winnings & " credits", "Congrats!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            ' adds the winning variable to the credits.
            credits = credits + Winnings
        End If

        ' procedure repeats for matches of all other 7 images.

        If image1 = 2 And image2 = 2 And image3 = 2 Then
            Winnings = 1000 * Bet
            MessageBox.Show("Congrats, you have won " & Winnings & " credits", "Congrats!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            credits = credits + Winnings
        End If

        If image1 = 3 And image2 = 3 And image3 = 3 Then
            Winnings = 850 * Bet
            MessageBox.Show("Congrats, you have won " & Winnings & " credits", "Congrats!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            credits = credits + Winnings
        End If

        If image1 = 4 And image2 = 4 And image3 = 4 Then
            Winnings = 700 * Bet
            MessageBox.Show("Congrats, you have won " & Winnings & " credits", "Congrats!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            credits = credits + Winnings
        End If

        If image1 = 5 And image2 = 5 And image3 = 5 Then
            Winnings = 500 * Bet
            MessageBox.Show("Congrats, you have won " & Winnings & " credits", "Congrats!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            credits = credits + Winnings
        End If

        If image1 = 6 And image2 = 6 And image3 = 6 Then
            Winnings = 400 * Bet
            MessageBox.Show("Congrats, you have won " & Winnings & " credits", "Congrats!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            credits = credits + Winnings
        End If

        If image1 = 72 And image2 = 7 And image3 = 7 Then
            Winnings = 250 * Bet
            MessageBox.Show("Congrats, you have won " & Winnings & " credits", "Congrats!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            credits = credits + Winnings
        End If

        If image1 = 8 And image2 = 8 And image3 = 8 Then
            Winnings = 100 * Bet
            MessageBox.Show("Congrats, you have won " & Winnings & " credits", "Congrats!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            credits = credits + Winnings
        End If

        ' displays the updated credits variable after a win.
        lblValTotalCred.Text = credits

    End Sub

    Private Sub tmrProgressive_Tick(sender As Object, e As EventArgs) Handles tmrProgressive.Tick

        ' everytime the timer ticks, 100 is added to the progressive variable.
        ' the number is big and the timer interval is fast to give a true random number to the progressive value. (so fast it is random)
        progressive += 100

        ' select case for the value of the progressive
        Select Case progressive
            ' if the progressive value gets to 5000, resets to 0. (5000 is the max)
            Case 5000
                progressive = 0
        End Select

        ' displays the progressive value in the value label.
        lblValProgressive.Text = progressive

    End Sub

    Private Sub frmPokeslots_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' starts a loop that displays an input box, and checks on the value of the variable.
        Do
            ' displays an input box that is for the inputcred variable
            inputcred = InputBox("Please enter the amount of money you would like to put into the machine.", "Please input your credits")

            ' if the variable is not numeric, gives an error message.
            If Not IsNumeric(inputcred) Then
                MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If

            ' if the variable is larger than 100,000, gives an error message
            If Val(inputcred) > 100001 Then
                MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If

            ' if the variable is 0 or smaller than (negative), gives an error message
            If Val(inputcred) < 0 Then
                MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If

            ' runs the loop until the inputcred variable is a numeric value between 0 and 100,001
        Loop Until Val(inputcred) > 0 And Val(inputcred) < 100001 And IsNumeric(inputcred)

        ' multiplies inputcred by 4 to give it a credit value.
        credits = inputcred * 4
        ' displays the credits.
        lblValTotalCred.Text = credits

    End Sub

    Private Sub radBet1_CheckedChanged(sender As Object, e As EventArgs) Handles radBet1.CheckedChanged

        ' when checked, sets the bet value to 1
        Bet = 1

    End Sub

    Private Sub radBet2_CheckedChanged(sender As Object, e As EventArgs) Handles radBet2.CheckedChanged

        ' when checked, sets the bet value to 2
        Bet = 2

    End Sub

    Private Sub radBet3_CheckedChanged(sender As Object, e As EventArgs) Handles radBet3.CheckedChanged

        ' when checked, sets the bet value to 3
        Bet = 3

    End Sub

    Private Sub btnCashIn_Click(sender As Object, e As EventArgs) Handles btnCashIn.Click

        ' starts a cashin loop similar to the form load loop. 
        Do
            ' starts an input box for the cashin variable
            cashin = InputBox("Please enter the amount of money you would like to put into the machine.", "Please input your credits")

            ' if cashin is not numeric, gives an error message
            If Not IsNumeric(cashin) Then
                MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If

            ' if cashin is over 100,000, gives an error message
            If Val(cashin) > 100001 Then
                MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If

            ' if cashin is less or = to 0, gives an error message
            If Val(cashin) < -1 Then
                MessageBox.Show("Please enter a positive number value under $100,000.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            End If

            ' loops until cashin is numeric, between 0 and 100,001
        Loop Until Val(cashin) > -1 And Val(cashin) < 100001 And IsNumeric(cashin)

        ' multiplies cashin by 4 to convert it to credits
        credits = credits + (cashin * 4)
        ' displays credit count.
        lblValTotalCred.Text = credits

    End Sub

    Private Sub btnCashout_Click(sender As Object, e As EventArgs) Handles btnCashout.Click

        ' displays a message box thanking you for playing, displaying your winnings
        MessageBox.Show("Thank you for playing! You won $ " & credits / 4 & ".", "Thank you!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

        ' ends program
        End

    End Sub
End Class
