using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using SoundForge;


// TODO: Need to handle situation where the total segment length is greater than the MOH length. 
// TODO: Consider adding an "Add" button to add segments to the list.
// TODO: Log file to keep track of previously opened directories and other specs.


public class Form1 : Form
{
    string segsFolder = String.Empty;
    string musicFilePath = String.Empty;
    string outPutDirectory = String.Empty;

    List<string> segsList = new List<string>();

    // Tracks our segments so we can deal with duplicated segments during our mixing action.
    Dictionary<string, string> segDictionary = new Dictionary<string, string>();
    // Tracks the number of each duplicated segment so we can keep them straight.
    Dictionary<string, int> segDupeDictionary = new Dictionary<string, int>();

    double totalLength = 0;
    double dBLevelDrop = 0;

    bool haveSegsFolder = false;
    private Button removeButton;
    private TextBox musicFileTextbox;
    private Button addButton;
    bool haveMusicFilePath = false;
    
    private void segsSelectButton_Click(object sender, EventArgs e)
    {
        if ((segsFolder = SfHelpers.ChooseDirectory("Choose the folder with your segments.", @"C:\")) != null)
        {
            haveSegsFolder = true;
            segsSelectButton.Text = "Segments: " + segsFolder;
        }
        else
        {
            return;
        }

        addSegments(Directory.GetFiles(segsFolder));

        
    }

    private void addSegments(string[] segs)
    {
        // Bool to keep track of any duplicate segments that user might try to load.
        bool duplicateKey = false;

        foreach (string s in segs)
        {
            // Handle attempts to add duplicate segments
            if (!listBox1.Items.Contains(s))
            {
                //dictionary uses the seg filename as a key and the path to the seg as a value
                segDictionary.Add(s, s);
                // Keeps track of the number of segment duplicates for naming purposes.
                segDupeDictionary.Add(s, 0);
                segsList.Add(s);
                listBox1.Items.Add(s);
            }
            else
            {
                duplicateKey = true;
            }

        }

        if (duplicateKey)
        {
            MessageBox.Show("One or more of these segments is already loaded. Please use the Duplicate button for " +
                        "segments you wish to repeat");
        }
    }
    

    private void musicSelectButton_Click(object sender, EventArgs e)
    {
        OpenFileDialog opener = new OpenFileDialog();
        opener.Filter = "WAV files (*.wav)|*.wav";
        opener.RestoreDirectory = true;
        opener.Title = "Choose the music bed file.";

        if ((musicFilePath = PathGetter(opener)) != "UserCancelled")
        {
            haveMusicFilePath = true;

            outPutDirectory = Directory.GetParent(musicFilePath).ToString();

            string music = Path.GetFileName(musicFilePath);
            musicFileTextbox.Text = music;
            musicFileTextbox.ForeColor = Color.Black;
        }

    }

    private void outFormatButton_Click(object sender, EventArgs e)
    {
                     

        //Choose our render as settings and filename.
        //Consider fancying this up a bit in the GUI.
    }

    private void upButton_Click(object sender, EventArgs e)
    {
        MoveListBoxItem(-1);
    }

    private void downButton_Click(object sender, EventArgs e)
    {
        MoveListBoxItem(1);
    }

    private void dupButton_Click(object sender, EventArgs e)
    {
        //duplicates a seg (or group of selected segs?) in our segment list
        if (listBox1.SelectedItem == null || listBox1.SelectedIndex < 0)
        {
            return;  //No item selected. Do nothin.
        }

        string selected = listBox1.SelectedItem.ToString();

        // Add 1 to our seg dupe dictionary to keep the names distinct.
        segDupeDictionary[selected] += 1; 

        string dupSeg = selected + " - copy" + segDupeDictionary[selected];
        //add the copied segment to our dictionary so it can be used during our mixing action
        segDictionary.Add(dupSeg, segDictionary[selected]);
        segDupeDictionary.Add(dupSeg, 0);
        listBox1.Items.Add(dupSeg);

    }

    public void MoveListBoxItem(int direction)
    {
        if (listBox1.SelectedItem == null || listBox1.SelectedIndex < 0)
        {
            return; // No item selected. Do nothing.
        }

        //calculate new index using move direction
        int newIndex = listBox1.SelectedIndex + direction;

        if (newIndex < 0 || newIndex >= listBox1.Items.Count)
        {
            return; // index out of range. nothing to do.
        }

        object selected = listBox1.SelectedItem;

        //removing removable element
        listBox1.Items.Remove(selected);
        //insert it in the new position
        listBox1.Items.Insert(newIndex, selected);
        //Restore selection
        listBox1.SetSelected(newIndex, true);
    }

    public void makeButton_Click(object sender, EventArgs e)
    {
        if (haveSegsFolder && haveMusicFilePath)
        {
            //parse our production length and verify it to continue
            if (TimeStringToSeconds(lengthTextBox.Text.ToString(), out totalLength))
            {
                //parse the duck level to a double to get our level drop at VO segments
                if (TextToDBLevel(duckLevelTextBox.Text.ToString(), out dBLevelDrop))
                {
                    //if we've made it this far, we're ready to do our mixing.

                    //pull our segments in order from the listBox
                    List<string> segments = new List<string>();
                    foreach (string item in listBox1.Items)
                    {
                        //Lookup the seg path and add it to our seg list
                        segments.Add(segDictionary[item]);
                    }
                    
                    //get the total length of the segments
                    double segsLength = FindSegsLength(segments, appl);
                    //get the length of silence between segments
                    double amtSilence = FindSilenceLength(totalLength, segsLength, segments.Count);

                    //Lay our music bed down in our output file
                    ISfFileHost musicFile = appl.OpenFile(musicFilePath, false, true);

                    if (musicFile.SampleRate != 44100)
                    {
                        //resample and channel convert music file to 44.1Khz mono
                        musicFile.DoResample(44100, 4, EffectOptions.WaitForDoneOrCancel | EffectOptions.EffectOnly);
                    }
                    if (musicFile.Channels > 1)
                    {
                        double[,] aGainMap = new double[1, 2] { { 0.5, 0.5 } };
                        musicFile.DoConvertChannels(1, 0, aGainMap, EffectOptions.WaitForDoneOrCancel | EffectOptions.EffectOnly);
                    }

                    SfAudioSelection music = new SfAudioSelection(musicFile);

                    long prodLength = musicFile.SecondsToPosition(totalLength);

                    ISfFileHost outFile = appl.NewFile(musicFile.DataFormat, false);

                    // Start our mixing operation.

                    // Paste our music file into the out file until it is longer than the total prod length.
                    while (outFile.Length < prodLength)
                    {
                        outFile.OverwriteAudio(outFile.Length, 0, musicFile, music);
                    }

                    musicFile.Close(CloseOptions.DiscardChanges);

                    //crop our outFile to the desired length.
                    outFile.CropAudio(0, prodLength);

                    //Normalize the music volume to a comfortable level.
                    int normLevel = -24;    //might replace with user selectable value
                    bool normRMS = true;    //replace with user selection?

                    appl.DoEffect("Normalize", GetNormPreset(appl, normLevel, normRMS), EffectOptions.EffectOnly | EffectOptions.WaitForDoneOrCancel);

                    //Get the start position for our mixing operation.
                    double startPosition = amtSilence / 2;

                    // Fade out music at the end of our file
                    double fadeLength = 2.0;
                    if (startPosition < 2)
                    {
                        fadeLength = startPosition / 2;
                    }
                    Int64 ccFade = outFile.Length - outFile.SecondsToPosition(fadeLength);
                    outFile.DoEffect("Graphic Fade", "-6dB exponential fade out", new SfAudioSelection(ccFade, outFile.Length), EffectOptions.EffectOnly);

                    long pasteSec = outFile.SecondsToPosition(startPosition);
                    SfAudioCrossfade cFade = new SfAudioCrossfade(outFile.SecondsToPosition(.175));
                    double level = SfHelpers.dBToRatio(dBLevelDrop);

                    // Begin mixing our segments in.
                    foreach (string file in segments)
                    {
                        ISfFileHost segment = appl.OpenFile(file, false, true);

                        //remove any markers from the segment and add marker with the filename
                        segment.Markers.Clear();
                        segment.Markers.AddMarker(0, Path.GetFileName(file));

                        // resample to 44100Hz and make sure the file is mono to match music
                        ResampleAndMono(segment, appl, 44100);

                        SfAudioSelection pasteSection = new SfAudioSelection(pasteSec, segment.Length);

                        outFile.DoMixReplace(pasteSection, level, 1, segment, SfAudioSelection.All, cFade, cFade, EffectOptions.EffectOnly | EffectOptions.WaitForDoneOrCancel);

                        pasteSec += (segment.Length + outFile.SecondsToPosition(amtSilence));
                        segment.Close(CloseOptions.DiscardChanges);
                    }

                    outFile.Save(SaveOptions.AlwaysPromptForFilename);
                    outFile.Close(CloseOptions.SaveChanges);
                }
            }
            else
            {
                MessageBox.Show("Your MOH Length is not formatted correctly. Please try again.");
            }
        }
        else
        {
            MessageBox.Show("Something is missing. Did you choose a Segments folder and a Music file?");
            return;
        }
    }

    public ISfFileHost ResampleAndMono(ISfFileHost inFile, IScriptableApp appl, uint rate)
    {
        inFile.DoResample(rate, 4, EffectOptions.WaitForDoneOrCancel | EffectOptions.EffectOnly);
        if (inFile.Channels > 1)
        {
            double[,] aGainMap = new double[1, 2] { { 0.5, 0.5 } };
            inFile.DoConvertChannels(1, 0, aGainMap, EffectOptions.WaitForDoneOrCancel | EffectOptions.EffectOnly);
        }

        return inFile;
    }

    public ISfGenericPreset GetNormPreset(IScriptableApp appl, int level, bool rms)
    {
        //takes in a integer level and a bool for rms and creates a norm preset
        //----Example call: ISfGenericPreset normNeg4 = GetNormPreset(app, -4, false);-----

        ISfGenericEffect fx = appl.FindEffect("Normalize");
        ISfGenericPreset preset0 = fx.GetPreset(0);
        byte[] abData = preset0.Bytes;
        string levelString = level.ToString();

        ISfGenericPreset normPreset = new SoundForge.SfGenericPreset(levelString, fx, abData);
        Fields_Normalize field1 = Fields_Normalize.FromPreset(normPreset);

        //If 0dB or a positive dB was requested, set the NormTo field as 0.
        if (level >= 0)
        {
            field1.NormalizeTo = 0;
        }
        else
            field1.NormalizeTo = SfHelpers.dBToRatio(level);

        field1.RMS = rms;
        if (rms)
        {
            field1.IfClip = Fields_Normalize.ClipAction.Compress;
        }

        field1.ToPreset(normPreset);

        return normPreset;
    }

    public bool TimeStringToSeconds(string timeString, out double oSeconds)
    {
        // Take a string representing mm:ss and parse it to seconds
        
        double totSeconds = 0;
        bool canParse = false;

        // Verify that the string is not null and that it has a :
        if (timeString != null && timeString.Contains(":"))
        {
            string[] timeSplit = timeString.Split(':');
            double result = 0;

            if (Double.TryParse(timeSplit[0], out result))
            {
                totSeconds += result * 60;
                if (Double.TryParse(timeSplit[1], out result))
                {
                    totSeconds += result;
                    canParse = true;
                }
            }
        }

        oSeconds = totSeconds;
        return canParse;
    }

    public bool TextToDBLevel(string dBString, out double levelDrop)
    {
        double result = 0;
        //return true if the input string can be parsed to a double
        if (Double.TryParse(dBString, out result))
        {
            levelDrop = result;
            return true;
        }
        //return false if the input string cannot be parsed to a double
        else
        {
            MessageBox.Show("Your Duck Level is not formatted correctly. Please try again.");
            levelDrop = result;
            return false;
        }
    }

    public double FindSegsLength(List<string> filePaths, IScriptableApp zApp)
    {
        double total = 0.000;
        foreach (string s in filePaths)
        {
            ISfFileHost file = zApp.OpenFile(s, true, true);
            total += file.PositionToSeconds(file.Length);
            file.Close(CloseOptions.DiscardChanges);
        }

        return total;
    }

    public double FindSilenceLength(double prodLength, double voLength, double segCount)
    {
        double silenceLength = (prodLength - voLength) / segCount;
        silenceLength = Math.Floor(silenceLength * 100) / 100;  //rounds down to 3 decimal places
        return silenceLength;
    }

    public String PathGetter(OpenFileDialog open)
    {
        String path = String.Empty;

        DialogResult result = open.ShowDialog();

        if (result == DialogResult.OK)
        {
            try
            {
                path = open.FileName;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Could not read file from disk. Original Error: " + ex.Message);
            }
        }
        else if (result == DialogResult.Cancel)
        {
            path = "UserCancelled";
        }

        return path;
    }

    //This method has to be public for SF to see the controls.
    public void InitializeComponent()
    {
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.upButton = new System.Windows.Forms.Button();
            this.downButton = new System.Windows.Forms.Button();
            this.segsSelectButton = new System.Windows.Forms.Button();
            this.musicSelectButton = new System.Windows.Forms.Button();
            this.makeButton = new System.Windows.Forms.Button();
            this.lengthTextBox = new System.Windows.Forms.TextBox();
            this.mohLengthLabel = new System.Windows.Forms.Label();
            this.duckLevelLabel = new System.Windows.Forms.Label();
            this.duckLevelTextBox = new System.Windows.Forms.TextBox();
            this.dupButton = new System.Windows.Forms.Button();
            this.removeButton = new System.Windows.Forms.Button();
            this.musicFileTextbox = new System.Windows.Forms.TextBox();
            this.addButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(13, 75);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(213, 173);
            this.listBox1.TabIndex = 0;
            // 
            // upButton
            // 
            this.upButton.Location = new System.Drawing.Point(232, 119);
            this.upButton.Name = "upButton";
            this.upButton.Size = new System.Drawing.Size(75, 23);
            this.upButton.TabIndex = 1;
            this.upButton.Text = "Move Up";
            this.upButton.UseVisualStyleBackColor = true;
            this.upButton.Click += new System.EventHandler(this.upButton_Click);
            // 
            // downButton
            // 
            this.downButton.Location = new System.Drawing.Point(232, 148);
            this.downButton.Name = "downButton";
            this.downButton.Size = new System.Drawing.Size(75, 23);
            this.downButton.TabIndex = 2;
            this.downButton.Text = "Move Down";
            this.downButton.UseVisualStyleBackColor = true;
            this.downButton.Click += new System.EventHandler(this.downButton_Click);
            // 
            // segsSelectButton
            // 
            this.segsSelectButton.Location = new System.Drawing.Point(13, 30);
            this.segsSelectButton.Name = "segsSelectButton";
            this.segsSelectButton.Size = new System.Drawing.Size(294, 30);
            this.segsSelectButton.TabIndex = 6;
            this.segsSelectButton.Text = "Choose Segments Folder";
            this.segsSelectButton.UseVisualStyleBackColor = true;
            this.segsSelectButton.Click += new System.EventHandler(this.segsSelectButton_Click);
            // 
            // musicSelectButton
            // 
            this.musicSelectButton.Location = new System.Drawing.Point(12, 266);
            this.musicSelectButton.Name = "musicSelectButton";
            this.musicSelectButton.Size = new System.Drawing.Size(296, 30);
            this.musicSelectButton.TabIndex = 7;
            this.musicSelectButton.Text = "Choose Music File";
            this.musicSelectButton.UseVisualStyleBackColor = true;
            this.musicSelectButton.Click += new System.EventHandler(this.musicSelectButton_Click);
            // 
            // makeButton
            // 
            this.makeButton.Location = new System.Drawing.Point(101, 391);
            this.makeButton.Name = "makeButton";
            this.makeButton.Size = new System.Drawing.Size(125, 32);
            this.makeButton.TabIndex = 9;
            this.makeButton.Text = "Make MOH";
            this.makeButton.UseVisualStyleBackColor = true;
            this.makeButton.Click += new System.EventHandler(this.makeButton_Click);
            // 
            // lengthTextBox
            // 
            this.lengthTextBox.Location = new System.Drawing.Point(89, 343);
            this.lengthTextBox.Name = "lengthTextBox";
            this.lengthTextBox.Size = new System.Drawing.Size(55, 20);
            this.lengthTextBox.TabIndex = 10;
            this.lengthTextBox.Text = "mm:ss";
            // 
            // mohLengthLabel
            // 
            this.mohLengthLabel.AutoSize = true;
            this.mohLengthLabel.Location = new System.Drawing.Point(12, 346);
            this.mohLengthLabel.Name = "mohLengthLabel";
            this.mohLengthLabel.Size = new System.Drawing.Size(71, 13);
            this.mohLengthLabel.TabIndex = 11;
            this.mohLengthLabel.Text = "MOH Length:";
            // 
            // duckLevelLabel
            // 
            this.duckLevelLabel.AutoSize = true;
            this.duckLevelLabel.Location = new System.Drawing.Point(164, 346);
            this.duckLevelLabel.Name = "duckLevelLabel";
            this.duckLevelLabel.Size = new System.Drawing.Size(65, 13);
            this.duckLevelLabel.TabIndex = 12;
            this.duckLevelLabel.Text = "Duck Level:";
            // 
            // duckLevelTextBox
            // 
            this.duckLevelTextBox.Location = new System.Drawing.Point(235, 343);
            this.duckLevelTextBox.Name = "duckLevelTextBox";
            this.duckLevelTextBox.Size = new System.Drawing.Size(72, 20);
            this.duckLevelTextBox.TabIndex = 13;
            this.duckLevelTextBox.Text = "-20";
            // 
            // dupButton
            // 
            this.dupButton.Location = new System.Drawing.Point(232, 177);
            this.dupButton.Name = "dupButton";
            this.dupButton.Size = new System.Drawing.Size(75, 23);
            this.dupButton.TabIndex = 14;
            this.dupButton.Text = "Duplicate";
            this.dupButton.UseVisualStyleBackColor = true;
            this.dupButton.Click += new System.EventHandler(this.dupButton_Click);
            // 
            // removeButton
            // 
            this.removeButton.Location = new System.Drawing.Point(233, 207);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(75, 23);
            this.removeButton.TabIndex = 15;
            this.removeButton.Text = "Remove";
            this.removeButton.UseVisualStyleBackColor = true;
            this.removeButton.Click += new System.EventHandler(this.removeButton_Click);
            // 
            // musicFileTextbox
            // 
            this.musicFileTextbox.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.musicFileTextbox.ForeColor = System.Drawing.SystemColors.InactiveCaption;
            this.musicFileTextbox.Location = new System.Drawing.Point(12, 303);
            this.musicFileTextbox.Name = "musicFileTextbox";
            this.musicFileTextbox.ReadOnly = true;
            this.musicFileTextbox.Size = new System.Drawing.Size(296, 20);
            this.musicFileTextbox.TabIndex = 16;
            this.musicFileTextbox.Text = "Music File Not Selected";
            // 
            // addButton
            // 
            this.addButton.Location = new System.Drawing.Point(233, 90);
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(75, 23);
            this.addButton.TabIndex = 17;
            this.addButton.Text = "Add";
            this.addButton.UseVisualStyleBackColor = true;
            this.addButton.Click += new System.EventHandler(this.addButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Wave Files(*.WAV)|*.WAV";
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Select Wave Files to Add";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(320, 445);
            this.Controls.Add(this.addButton);
            this.Controls.Add(this.musicFileTextbox);
            this.Controls.Add(this.removeButton);
            this.Controls.Add(this.dupButton);
            this.Controls.Add(this.duckLevelTextBox);
            this.Controls.Add(this.duckLevelLabel);
            this.Controls.Add(this.mohLengthLabel);
            this.Controls.Add(this.lengthTextBox);
            this.Controls.Add(this.makeButton);
            this.Controls.Add(this.musicSelectButton);
            this.Controls.Add(this.segsSelectButton);
            this.Controls.Add(this.downButton);
            this.Controls.Add(this.upButton);
            this.Controls.Add(this.listBox1);
            this.Name = "Form1";
            this.Text = "MOH Maker for Sound Forge";
            this.ResumeLayout(false);
            this.PerformLayout();

    }
    
    private IScriptableApp appl;
    private System.Windows.Forms.ListBox listBox1;
    private System.Windows.Forms.Button upButton;
    private System.Windows.Forms.Button downButton;
    private System.Windows.Forms.Button segsSelectButton;
    private System.Windows.Forms.Button musicSelectButton;
    private System.Windows.Forms.Button makeButton;
    private System.Windows.Forms.TextBox lengthTextBox;
    private System.Windows.Forms.Label mohLengthLabel;
    private System.Windows.Forms.Label duckLevelLabel;
    private System.Windows.Forms.TextBox duckLevelTextBox;
    private System.Windows.Forms.Button dupButton;
    private System.Windows.Forms.OpenFileDialog openFileDialog1;

    public IScriptableApp Appl
    {
        get { return this.appl; }
        set { this.appl = value; }
    }

    private void removeButton_Click(object sender, EventArgs e)
    {
        if (listBox1.SelectedItem == null || listBox1.SelectedIndex < 0)
        {
            return;  // No selection. Do nothing.
        }

        object toRemove = listBox1.SelectedItem;
        // Remove segment from our dictionary.
        segDictionary.Remove(toRemove.ToString());
        listBox1.Items.Remove(toRemove);
    }

    private void addButton_Click(object sender, EventArgs e)
    {
        DialogResult dr = this.openFileDialog1.ShowDialog();
        if (dr == System.Windows.Forms.DialogResult.OK)
        {
            addSegments(openFileDialog1.FileNames);
        }
    }
}

public class EntryPoint
{
    public void Begin(IScriptableApp app)
    {
        Form1 theForm = new Form1();
        theForm.InitializeComponent();
        theForm.Appl = app;
        theForm.ShowDialog();
    }


    public void FromSoundForge(IScriptableApp app)
    {
        ForgeApp = app; //execution begins here
        app.SetStatusText(String.Format("Script '{0}' is running.", Script.Name));
        Begin(app);
        app.SetStatusText(String.Format("Script '{0}' is done.", Script.Name));
    }
    public static IScriptableApp ForgeApp = null;
    public static void DPF(string sz) { ForgeApp.OutputText(sz); }
    public static void DPF(string fmt, params object[] args) { ForgeApp.OutputText(String.Format(fmt, args)); }
} //EntryPoint



