using System;
using ZOSAPI;
using ZOSAPI.Analysis;
using ZOSAPI.Analysis.Data;
using ZOSAPI.Analysis.Settings;

namespace CSharpUserOperandApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            // Find the installed version of OpticStudio
            bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            // Note -- uncomment the following line to use a custom initialization path
            //bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(@"C:\Program Files\OpticStudio\");
            if (isInitialized)
            {
                LogInfo("Found OpticStudio at: " + ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory());
            }
            else
            {
                HandleError("Failed to locate OpticStudio!");
                return;
            }

            BeginUserOperand();
        }

        static void BeginUserOperand()
        {
            // Create the initial connection class
            ZOSAPI_Connection TheConnection = new ZOSAPI_Connection();

            // Attempt to connect to the existing OpticStudio instance
            IZOSAPI_Application TheApplication = null;
            try
            {
                TheApplication = TheConnection.ConnectToApplication(); // this will throw an exception if not launched from OpticStudio
            }
            catch (Exception ex)
            {
                HandleError(ex.Message);
                return;
            }
            if (TheApplication == null)
            {
                HandleError("An unknown connection error occurred!");
                return;
            }

            // Check the connection status
            if (!TheApplication.IsValidLicenseForAPI)
            {
                HandleError("Failed to connect to OpticStudio: " + TheApplication.LicenseStatus);
                return;
            }
            if (TheApplication.Mode != ZOSAPI_Mode.Operand)
            {
                HandleError("User plugin was started in the wrong mode: expected Operand, found " + TheApplication.Mode.ToString());
                return;
            }

            // Read the operand arguments
            double Hx = TheApplication.OperandArgument1;
            double Hy = TheApplication.OperandArgument2;
            double Px = TheApplication.OperandArgument3;
            double Py = TheApplication.OperandArgument4;

            // Initialize the output array
            int maxResultLength = TheApplication.OperandResults.Length;
            double[] operandResults = new double[maxResultLength];

            IOpticalSystem TheSystem = TheApplication.PrimarySystem;

            // Add your custom code here...

            // Open the Wavefront Map
            IA_ TheWavefrontMap = TheSystem.Analyses.New_WavefrontMap();

            // Change settings here ...
            // (Hx, Hy, Px, and Py can be used as user inputs)
            IAS_WavefrontMap TheWavefrontMapSettings = (IAS_WavefrontMap)TheWavefrontMap.GetSettings();

            // As an example, I'm showing how to change the field through the Hx parameter
            TheWavefrontMapSettings.Field.SetFieldNumber((int)Hx);

            // Run the analysis
            TheWavefrontMap.ApplyAndWaitForCompletion();

            // Get results
            IAR_ TheWavefrontMapResults = TheWavefrontMap.GetResults();

            // Read data, and compute peak-to-valley
            try
            {
                // Initialization
                int XRange, YRange;
                double valley = double.PositiveInfinity;
                double peak = double.NegativeInfinity;

                // Grid size
                XRange = (int)TheWavefrontMapResults.GetDataGrid(0).Nx;
                YRange = (int)TheWavefrontMapResults.GetDataGrid(0).Ny;

                // Grid data
                double[,] WavefrontData = TheWavefrontMapResults.GetDataGrid(0).Values;

                // Process grid data (in this case, extract minimum, and maximum to compute the peak-to-valley)
                for (int XX = 0; XX < XRange; XX++)
                {
                    for (int YY = 0; YY < YRange; YY++)
                    {
                        if (WavefrontData[XX, YY] > peak)
                        {
                            peak = WavefrontData[XX, YY];
                        }
                        if (WavefrontData[XX, YY] < valley)
                        {
                            valley = WavefrontData[XX, YY];
                        }
                    }
                }

                // Return peak-to-valley
                operandResults[0] = Math.Abs(peak-valley);
            }
            catch
            {
                operandResults[0] = -1;
            }

            // Clean up
            FinishUserOperand(TheApplication, operandResults);
        }

        static void FinishUserOperand(IZOSAPI_Application TheApplication, double[] resultData)
        {
            // Note - OpticStudio will wait for the operand to complete until this application exits 
            if (TheApplication != null)
            {
                TheApplication.OperandResults.WriteData(resultData.Length, resultData);
            }
        }

        static void LogInfo(string message)
        {
            // TODO - add custom logging
            Console.WriteLine(message);
        }

        static void HandleError(string errorMessage)
        {
            // TODO - add custom error handling
            throw new Exception(errorMessage);
        }

    }
}
