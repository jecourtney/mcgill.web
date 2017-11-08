using System;
using System.Collections.Generic;
using System.IO;
using McGill.Library;

namespace McGill.Web
{
    public static class AirFlowTechTools
    {
        private static List<string[]> _RectData = null;
        private static List<string[]> _AngleData = null;

        #region VFR
        /// <summary>
        /// Calculates the volume flow rate.
        /// </summary>
        /// <param name="mInput">The m input.</param>
        /// <param name="bToCFM">if set to <c>true</c> [b to CFM].</param>
        /// <returns></returns>
        public static decimal CalcVolumeFlowRate(decimal mInput, bool bToCFM)
        {
            decimal mOutput = 0M;

            if (bToCFM)
            {
                mOutput = UMCLib.Round(mInput * 2.11888M, 4);
            }
            else
            {
                mOutput = UMCLib.Round(mInput * 0.4719474M, 4);
            }

            return mOutput;
        }
        #endregion VFR

        #region CalcConversion
        /// <summary>
        /// Calculates the conversion.
        /// </summary>
        /// <param name="sCalcType">Type of the s calculate.</param>
        /// <param name="mRectMinor">The m rect minor.</param>
        /// <param name="mRectMajor">The m rect major.</param>
        /// <param name="mOvalMinor">The m oval minor.</param>
        /// <param name="mOvalMajor">The m oval major.</param>
        /// <param name="mDiamter">The m diamter.</param>
        /// <param name="mMinor">The m minor.</param>
        /// <param name="mResult1">The m result1.</param>
        /// <param name="mResult2">The m result2.</param>
        public static void CalcConversion(string sCalcType, decimal mRectMinor, decimal mRectMajor, decimal mOvalMinor, decimal mOvalMajor, decimal mDiamter, decimal mMinor,
            out decimal mResult1, out decimal mResult2)
        {
            mResult1 = 0M;
            mResult2 = 0M;

            if (sCalcType == "Rect")
            {
                decimal mEquivRect = CalcEquivRect(mRectMinor, mRectMajor);
                mOvalMajor = mOvalMinor + 1M;
                mDiamter = CalcEquivOval(mOvalMinor, mOvalMajor);

                while (Math.Abs(mDiamter - mEquivRect) >= 0.02M)
                {
                    if (mDiamter < mEquivRect)
                    {
                        mOvalMajor += 0.01M;
                    }
                    else if (mDiamter > mEquivRect)
                    {
                        mOvalMajor -= 0.01M;
                    }

                    mDiamter = CalcEquivOval(mOvalMinor, mOvalMajor);
                }

                mResult1 = UMCLib.Round(mDiamter, 2);
                mResult2 = UMCLib.Round(mOvalMajor, 2);
            }
            else if (sCalcType == "Oval")
            {
                decimal mEquivOval = CalcEquivOval(mOvalMinor, mOvalMajor);
                mRectMajor = mRectMinor + 1M;
                mDiamter = CalcEquivRect(mRectMinor, mRectMajor);

                while (Math.Abs(mDiamter - mEquivOval) >= 0.02M)
                {
                    if (mDiamter < mEquivOval)
                    {
                        mRectMajor += 0.01M;
                    }
                    else if (mDiamter > mEquivOval)
                    {
                        mRectMajor -= 0.01M;
                    }

                    mDiamter = CalcEquivRect(mRectMinor, mRectMajor);
                }

                mResult1 = UMCLib.Round(mDiamter, 2);
                mResult2 = UMCLib.Round(mRectMajor, 2);
            }
            else if (sCalcType == "Round")
            {
                mRectMinor = mMinor;
                mRectMajor = mMinor + 0.1M;

                decimal mEquivRect = CalcEquivRect(mRectMinor, mRectMajor);

                while (Math.Abs(mDiamter - mEquivRect) >= 0.02M)
                {
                    if (mDiamter > mEquivRect)
                    {
                        mRectMajor += 0.01M;
                    }
                    else if (mDiamter < mEquivRect)
                    {
                        mRectMajor -= 0.01M;
                    }

                    mEquivRect = CalcEquivRect(mRectMinor, mRectMajor);
                }

                mOvalMinor = mMinor;
                mOvalMajor = mMinor + 0.1M;

                decimal mEquivOval = CalcEquivOval(mOvalMinor, mOvalMajor);

                while (Math.Abs(mDiamter - mEquivOval) >= 0.02M)
                {
                    if (mDiamter > mEquivOval)
                    {
                        mOvalMajor += 0.01M;
                    }
                    else if (mDiamter < mEquivOval)
                    {
                        mOvalMajor -= 0.01M;
                    }

                    mEquivOval = CalcEquivOval(mOvalMinor, mOvalMajor);
                }

                mResult1 = UMCLib.Round(mRectMajor, 2);
                mResult2 = UMCLib.Round(mOvalMajor, 2);
            }
        }

        private static decimal CalcEquivOval(decimal mOvalMinor, decimal mOvalMajor)
        {
            decimal mPrimeter = CalcPrimeter(mOvalMinor, mOvalMajor);
            return 1.55M * UMCLib.Pow((mOvalMajor - mOvalMinor) * mOvalMinor + 0.25M * 3.14159M * UMCLib.Pow(mOvalMinor, 2M), 0.625M) / UMCLib.Pow(mPrimeter, 0.25M);
        }

        private static decimal CalcPrimeter(decimal mMinor, decimal mMajor)
        {
            return 3.14159M * mMinor + 2M * (mMajor - mMinor);
        }

        private static decimal CalcEquivRect(decimal mRectMinor, decimal mRectMajor)
        {
            return (1.30M * UMCLib.Pow(mRectMinor * mRectMajor, 0.625M)) / UMCLib.Pow(mRectMinor + mRectMajor, 0.25M);
        }
        #endregion CalcConversion

        #region Pressure
        /// <summary>
        /// Calculates the operating pressure.
        /// </summary>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mStiffenerSpacing">The m stiffener spacing.</param>
        /// <param name="mDuctTemp">The m duct temporary.</param>
        /// <param name="sStiffenerSize">Size of the s stiffener.</param>
        /// <returns></returns>
        /// <exception cref="Exception">Invalid combination of material, gauge, and duct construction.</exception>
        /// <exception cref="System.Exception">Unrecogined gauge. Valid gauges are 26, 24, 22, 20, 18, 16
        /// or
        /// Unrecogined gauge. Valid gauges are 26, 24, 22, 20, 18, 16
        /// or
        /// Unrecogined gauge. Valid gauges are 26, 24, 22, 20, 18, 16
        /// or
        /// Unrecogined material. Valid materials are Steel, Stainless Steel, Aluminum</exception>
        public static decimal CalcOperatingPressure(string sMaterial, int nGauge, bool bSpiral, decimal mDiameter, decimal mStiffenerSpacing, decimal mDuctTemp,
            out string sStiffenerSize)
        {
            if (!ValidGauge(sMaterial, nGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, gauge, and duct construction.");
            }

            decimal mModulus = 0M;
            decimal mThickness = 0M;
            decimal mUnitWeight = 0M;
            decimal mYieldStrength = 0M;
            decimal mDensity = 0M;

            MetalProperties(sMaterial, nGauge, out mModulus, out mThickness, out mUnitWeight, out mYieldStrength, out mDensity);
            mModulus = AdjustModulusToTemp(mModulus, mDuctTemp, false);

            decimal mPressure = 0M;

            if (bSpiral)
            {
                mPressure = (1M / 3M) * 41.3817M * mModulus * (UMCLib.Pow(mThickness, 3.01212M) / (UMCLib.Pow(mDiameter, 1.53827M) * UMCLib.Pow(mStiffenerSpacing, 0.77924M)));
            }
            else
            {
                mPressure = 8.304M * UMCLib.Pow(10M, 9M) * (mDiameter / mStiffenerSpacing) * UMCLib.Pow(mThickness / mDiameter, 2.5M) * (mModulus / 30000000M) * (1M / (52M + mDiameter) * (1M / 3.5M));
            }

            sStiffenerSize = StiffenerSize(mDiameter, mPressure, mStiffenerSpacing, mModulus);

            return UMCLib.Round(mPressure, 1);
        }

        /// <summary>
        /// Calculates the stiffener spacing.
        /// </summary>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mPressure">The m pressure.</param>
        /// <param name="mDuctTemp">The m duct temporary.</param>
        /// <param name="sStiffenerSize">Size of the s stiffener.</param>
        /// <returns></returns>
        /// <exception cref="Exception">Invalid combination of material, gauge, and duct construction.</exception>
        /// <exception cref="System.Exception">Invalid combination of material, gauge, and duct construction.</exception>
        public static decimal CalcStiffenerSpacing(string sMaterial, int nGauge, bool bSpiral, decimal mDiameter, decimal mPressure, decimal mDuctTemp,
            out string sStiffenerSize)
        {
            if (!ValidGauge(sMaterial, nGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, gauge, and duct construction.");
            }

            decimal mModulus = 0M;
            decimal mThickness = 0M;
            decimal mUnitWeight = 0M;
            decimal mYieldStrength = 0M;
            decimal mDensity = 0M;

            MetalProperties(sMaterial, nGauge, out mModulus, out mThickness, out mUnitWeight, out mYieldStrength, out mDensity);
            mModulus = AdjustModulusToTemp(mModulus, mDuctTemp, false);
            mPressure = Math.Abs(mPressure);

            decimal mStiffenerSpacing = 0M;

            if (bSpiral)
            {
                mStiffenerSpacing = UMCLib.Pow((41.3817M * mModulus * UMCLib.Pow(mThickness, 3.01212M)) / (mPressure * 3M * UMCLib.Pow(mDiameter, 1.53827M)), 1M / 0.77924M);
            }
            else
            {
                mStiffenerSpacing = 8.304M * UMCLib.Pow(10M, 9M) * (mDiameter / mPressure) * UMCLib.Pow(mThickness / mDiameter, 2.5M) * (mModulus / (30M * UMCLib.Pow(10M, 6M))) / (52M + mDiameter) * (1M / 3.5M);
            }

            sStiffenerSize = StiffenerSize(mDiameter, mPressure, mStiffenerSpacing, mModulus);

            return UMCLib.Round(mStiffenerSpacing, 1);
        }

        /// <summary>
        /// Calculates the minimum thickness.
        /// </summary>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mStiffenerSpacing">The m stiffener spacing.</param>
        /// <param name="mPressure">The m pressure.</param>
        /// <param name="mDuctTemp">The m duct temporary.</param>
        /// <param name="sClass">The s class.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="sStiffenerSize">Size of the s stiffener.</param>
        /// <returns></returns>
        public static decimal CalcMinThickness(string sMaterial, bool bSpiral, decimal mDiameter, decimal mStiffenerSpacing, decimal mPressure, decimal mDuctTemp, string sClass,
            out int nGauge, out string sStiffenerSize)
        {
            decimal mModulus = 0M;
            decimal mThickness = 0M;
            decimal mUnitWeight = 0M;
            decimal mYieldStrength = 0M;
            decimal mDensity = 0M;

            MetalProperties(sMaterial, 20, out mModulus, out mThickness, out mUnitWeight, out mYieldStrength, out mDensity);
            mModulus = AdjustModulusToTemp(mModulus, mDuctTemp, false);
            mPressure = Math.Abs(mPressure);

            if (bSpiral)
            {
                mThickness = UMCLib.Pow((mPressure * 3M / (41.3817M * mModulus)) * UMCLib.Pow(mDiameter, 1.53827M) * UMCLib.Pow(mStiffenerSpacing, 0.77924M), 1M / 3.01212M) + ClassThickness(sClass);
            }
            else
            {
                mThickness = mDiameter * UMCLib.Pow(mPressure * (52M + mDiameter) * mStiffenerSpacing / (mDiameter * 2.373M * UMCLib.Pow(10M, 9M)), 1M / 2.5M) + ClassThickness(sClass);
            }

            nGauge = ThicknessToGauge(sMaterial, mThickness);
            sStiffenerSize = StiffenerSize(mDiameter, mPressure, mStiffenerSpacing, mModulus);

            return mThickness;
        }

        /// <summary>
        /// Calculates the negative pressure no stiffeners.
        /// </summary>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mDuctTemp">The m duct temporary.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Invalid combination of material, gauge, and duct construction.</exception>
        public static decimal CalcNegativePressureNoStiffeners(string sMaterial, int nGauge, bool bSpiral, decimal mDiameter, decimal mDuctTemp)
        {
            if (!ValidGauge(sMaterial, nGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, gauge, and duct construction.");
            }

            decimal mModulus = 0M;
            decimal mThickness = 0M;
            decimal mUnitWeight = 0M;
            decimal mYieldStrength = 0M;
            decimal mDensity = 0M;

            MetalProperties(sMaterial, nGauge, out mModulus, out mThickness, out mUnitWeight, out mYieldStrength, out mDensity);
            mModulus = AdjustModulusToTemp(mModulus, mDuctTemp, false);

            decimal mPressure = 0M;
            decimal mTD = mThickness / mDiameter;

            if (bSpiral)
            {
                mModulus = mModulus / 30E6M;

                switch (nGauge)
                {
                    case 28:
                        mPressure = mModulus * (-0.573M + 6724.6839M * mTD + 11255725M * UMCLib.Pow(mTD, 2M)) / 3M;
                        break;
                    case 26:
                        mPressure = mModulus * (UMCLib.Exp(-0.9427M + 115.0434M * UMCLib.Sqrt(mTD)) / 3M);
                        break;
                    case 24:
                        mPressure = mModulus * (0.30673179M + 30361.872M * mTD - 37646859M * UMCLib.Pow(mTD, 2M) + 1.9909081E10M * UMCLib.Pow(mTD, 3M)) / 3M;
                        break;
                    case 22:
                        mPressure = mModulus * (-16.494248M + 64625.103M * mTD - 57251360M * UMCLib.Pow(mTD, 2M) + 2.5400559E10M * UMCLib.Pow(mTD, 3M)) / 3M;
                        break;
                    case 20:
                        mPressure = mModulus * (14.47046M - 10242.539M * mTD - 2127069.9M * UMCLib.Pow(mTD, 2M) + 1.3513E10M * UMCLib.Pow(mTD, 3M)) / 3M;
                        break;
                    case 18:
                        mPressure = mModulus * (-14.929715M + 58219.341M * mTD - 45366164M * UMCLib.Pow(mTD, 2M) + 1.8968003E10M * UMCLib.Pow(mTD, 3M)) / 3M;
                        break;
                    case 16:
                    case 14:
                        mPressure = mModulus * (-14.215146M + 87239.75M * mTD - 76484166M * UMCLib.Pow(mTD, 2M) + 2.8124138E10M * UMCLib.Pow(mTD, 3M)) / 3M;
                        break;
                }
            }
            else
            {
                mPressure = 852.3M * mModulus * UMCLib.Pow(mThickness / mDiameter, 3M) / (52M + mDiameter);
            }

            return UMCLib.Round(mPressure, 1);
        }

        /// <summary>
        /// Calculates the burst pressure.
        /// </summary>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mDuctTemp">The m duct temporary.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Invalid combination of material, gauge, and duct construction.</exception>
        public static decimal CalcBurstPressure(string sMaterial, int nGauge, bool bSpiral, decimal mDiameter, decimal mDuctTemp)
        {
            if (!ValidGauge(sMaterial, nGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, gauge, and duct construction.");
            }

            decimal mModulus = 0M;
            decimal mThickness = 0M;
            decimal mUnitWeight = 0M;
            decimal mYieldStrength = 0M;
            decimal mDensity = 0M;

            MetalProperties(sMaterial, nGauge, out mModulus, out mThickness, out mUnitWeight, out mYieldStrength, out mDensity);
            mModulus = AdjustModulusToTemp(mModulus, mDuctTemp, false);

            decimal output = 0.7480847M * mModulus * UMCLib.Pow(mThickness / mDiameter, 1.4792M);

            return UMCLib.Round(output, 1);
        }

        #endregion Pressure

        #region Support
        /// <summary>
        /// Calculates the support.
        /// </summary>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mStiffenerSpacing">The m stiffener spacing.</param>
        /// <param name="mLoad">The m load.</param>
        /// <param name="mAirDensity">The m density.</param>
        /// <param name="mWind">The m wind.</param>
        /// <param name="mSnow">The m snow.</param>
        /// <param name="mSafetyFactor">The m safety factor.</param>
        /// <param name="mMaterialLoad">The m material load.</param>
        /// <param name="mAllowedDeflection">The m allowed deflection.</param>
        /// <param name="mDeflection">The m deflection.</param>
        /// <param name="mMaxLength">Maximum length of the m.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Invalid combination of material, gauge, and duct construction.</exception>
        public static string CalcSupport(string sMaterial, int nGauge, bool bSpiral, decimal mDiameter, decimal mStiffenerSpacing, decimal mLoad, decimal mAirDensity,
            decimal mWind, decimal mSnow, decimal mSafetyFactor, out decimal mMaterialLoad, out decimal mAllowedDeflection, out decimal mDeflection, out decimal mMaxLength)
        {
            if (!ValidGauge(sMaterial, nGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, gauge, and duct construction.");
            }

            decimal mModulus = 0M;
            decimal mThickness = 0M;
            decimal mUnitWeight = 0M;
            decimal mYieldStrength = 0M;
            decimal mMetalDensity = 0M;

            MetalProperties(sMaterial, nGauge, out mModulus, out mThickness, out mUnitWeight, out mYieldStrength, out mMetalDensity);

            mMaterialLoad = mLoad * (3.14159M * UMCLib.Pow(mDiameter, 2M) / 576M) * mAirDensity / 12M;
            decimal mDeadWeight = 0M;

            if (bSpiral)
            {
                if (mDiameter < 8M || mDiameter > 60M)
                {
                    mDeadWeight = (mSnow / 144M) * mDiameter + mUnitWeight * ((12M * 3.14159M * mDiameter / (5.394M - 0.75M)) * (5.394M / 1728M)) + mMaterialLoad;
                }
                else
                {
                    mDeadWeight = (mSnow / 144M) * mDiameter + mUnitWeight * ((12M * 3.14159M * mDiameter / (6.875M - 0.75M)) * (6.875M / 1728M)) + mMaterialLoad;
                }
            }
            else
            {
                mDeadWeight = (mSnow / 144M) * mDiameter + mUnitWeight * 3.14159M * mDiameter / 144M + mMaterialLoad;
            }

            decimal mWindVP = UMCLib.Pow((mWind * 5280M / 60M) / 4005M, 2) / 27.7M;
            decimal mLoadPerLength = mDeadWeight;

            if (mWindVP * mDiameter > mDeadWeight)
            {
                mLoadPerLength = mWindVP * mDiameter;
            }

            decimal mMom = MomentOfInertia(bSpiral, mDiameter, mThickness);

            SupportDeflection(mDiameter, mLoadPerLength, mStiffenerSpacing, mModulus, mMom, mSafetyFactor, out mDeflection, out mAllowedDeflection, out mMaxLength);

            mMaterialLoad = UMCLib.Round(mMaterialLoad, 4);

            if (mAllowedDeflection < mDeflection)
            {
                return "Fail";
            }
            else
            {
                return "Pass";
            }
        }


        /// <summary>
        /// Calculates the support.
        /// </summary>
        /// <param name="mInsulThickness">The m insul thickness.</param>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="nInnerGauge">The n inner gauge.</param>
        /// <param name="nOuterGauge">The n outer gauge.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mStiffenerSpacing">The m stiffener spacing.</param>
        /// <param name="mLoad">The m load.</param>
        /// <param name="mAirDensity">The m density.</param>
        /// <param name="mWind">The m wind.</param>
        /// <param name="mSnow">The m snow.</param>
        /// <param name="mSafetyFactor">The m safety factor.</param>
        /// <param name="mMaterialLoad">The m material load.</param>
        /// <param name="mAllowedDeflection">The m allowed deflection.</param>
        /// <param name="mDeflection">The m deflection.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Invalid combination of material, inner gauge, and duct construction.
        /// or
        /// Invalid combination of material, outer gauge, and duct construction.</exception>
        public static string CalcSupport(decimal mInsulThickness, string sMaterial, int nInnerGauge, int nOuterGauge, bool bSpiral, decimal mDiameter, decimal mStiffenerSpacing, decimal mLoad, decimal mAirDensity,
            decimal mWind, decimal mSnow, decimal mSafetyFactor, out decimal mMaterialLoad, out decimal mAllowedDeflection, out decimal mDeflection, out decimal mMaxLength)
        {
            if (!ValidGauge(sMaterial, nInnerGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, inner gauge, and duct construction.");
            }

            if (!ValidGauge(sMaterial, nOuterGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, outer gauge, and duct construction.");
            }

            decimal mModulus = 0M;
            decimal mInnerThickness = 0M;
            decimal mInnerUnitWeight = 0M;
            decimal mInnerYieldStrength = 0M;
            decimal mInnerDensity = 0M;
            decimal mOuterThickness = 0M;
            decimal mOuterUnitWeight = 0M;
            decimal mOuterYieldStrength = 0M;
            decimal mOuterDensity = 0M;

            MetalProperties(sMaterial, nInnerGauge, out mModulus, out mInnerThickness, out mInnerUnitWeight, out mInnerYieldStrength, out mInnerDensity);
            MetalProperties(sMaterial, nOuterGauge, out mModulus, out mOuterThickness, out mOuterUnitWeight, out mOuterYieldStrength, out mOuterDensity);

            decimal mOuterDiameter = mDiameter + 2M * mInsulThickness;

            mMaterialLoad = mLoad * (3.14159M * UMCLib.Pow(mDiameter, 2M) / 576M) * mAirDensity / 12M;
            decimal mDeadWeight = 0M;
            decimal mInsulDensity = 1M;

            if (bSpiral)
            {
                decimal mUnitWeightTotal = 0M;

                if (mDiameter < 8M || mDiameter > 60M)
                {
                    mUnitWeightTotal = mInnerUnitWeight * ((12M * 3.14159M * mDiameter / (5.394M - 0.75M)) * (5.394M / 1728M)) + mOuterUnitWeight * ((12M * 3.14159M * mOuterDiameter / (5.394M - 0.75M)) * (5.394M / 1728M)) + mMaterialLoad;
                }
                else
                {
                    mUnitWeightTotal = mInnerUnitWeight * ((12M * 3.14159M * mOuterDiameter / (6.875M - 0.75M)) * (6.875M / 1728M)) + mOuterUnitWeight * ((12M * 3.14159M * mOuterDiameter / (6.875M - 0.75M)) * (6.875M / 1728M)) + mMaterialLoad;
                }

                mDeadWeight = (mInsulDensity * 3.14159M * (UMCLib.Pow(mOuterDiameter / 2M, 2M) - UMCLib.Pow(mDiameter / 2, 2M)) / 144M) + (mSnow / 144M) * mOuterDiameter + mUnitWeightTotal;
            }
            else
            {
                decimal mUnitWeightTotal = mInnerUnitWeight * 3.14159M * mDiameter / 144M + mOuterUnitWeight * 3.14159M * mOuterDiameter / 144M + mMaterialLoad;
                mDeadWeight = (mInsulDensity * 3.14159M * (UMCLib.Pow(mOuterDiameter / 2M, 2M) - UMCLib.Pow(mDiameter / 2, 2M)) / 144M) + (mSnow / 144M) * mOuterDiameter + mUnitWeightTotal * 3.14159M * mOuterDiameter / 144M;
            }

            decimal mWindVP = UMCLib.Pow((mWind * 5280M / 60M) / 4005M, 2) / 27.7M;
            decimal mLoadPerLength = mDeadWeight;

            if (mWindVP * mOuterDiameter > mDeadWeight)
            {
                mLoadPerLength = mWindVP * mOuterDiameter;
            }

            decimal mMom = MomentOfInertia(bSpiral, mOuterDiameter, mOuterThickness);

            SupportDeflection(mOuterDiameter, mLoadPerLength, mStiffenerSpacing, mModulus, mMom, mSafetyFactor, out mDeflection, out mAllowedDeflection, out mMaxLength);

            mMaterialLoad = UMCLib.Round(mMaterialLoad, 4);

            if (mAllowedDeflection < mDeflection)
            {
                return "Fail";
            }
            else
            {
                return "Pass";
            }
        }

        #endregion Support

        #region Stack
        /// <summary>
        /// Calculates the stack.
        /// </summary>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mWind">The m wind.</param>
        /// <param name="mSafetyFactor">The m safety factor.</param>
        /// <param name="mHeight">Height of the m.</param>
        /// <param name="mCriticalVelocity">The m critical velocity.</param>
        /// <param name="mActualStress">The m actual stress.</param>
        /// <param name="mCriticalBuckling">The m critical buckling.</param>
        /// <param name="mCriticalYield">The m critical yield.</param>
        /// <param name="mDeflection">The m deflection.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Invalid combination of material, gauge, and duct construction.</exception>
        public static string CalcStack(string sMaterial, int nGauge, bool bSpiral, decimal mDiameter, decimal mWind, decimal mSafetyFactor, decimal mHeight,
            out decimal mCriticalVelocity, out decimal mActualStress, out decimal mCriticalBuckling, out decimal mCriticalYield, out decimal mDeflection)
        {
            if (!ValidGauge(sMaterial, nGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, gauge, and duct construction.");
            }

            decimal mModulus = 0M;
            decimal mThickness = 0M;
            decimal mUnitWeight = 0M;
            decimal mYieldStrength = 0M;
            decimal mDensity = 0M;

            MetalProperties(sMaterial, nGauge, out mModulus, out mThickness, out mUnitWeight, out mYieldStrength, out mDensity);

            decimal mMom = MomentOfInertia(bSpiral, mDiameter, mThickness);

            decimal mFreq = 3.9M * (mDiameter / UMCLib.Pow(mHeight * 12M, 2M)) * UMCLib.Pow(mModulus / mDensity, 0.5M);
            mCriticalVelocity = 0.284M * mDiameter * mFreq;

            decimal mWindVP = UMCLib.Pow((mWind * 5280M / 60M) / 4005M, 2) / 27.7M;
            decimal mLoadPerLength = mWindVP * mDiameter;

            decimal mMimimumThickness;
            StackDeflection(mDiameter, mLoadPerLength, mHeight, mModulus, mMom, out mDeflection, out mMimimumThickness);

            int nMinGauge = ThicknessToGauge(sMaterial, mMimimumThickness);

            mCriticalBuckling = 1.2M * mModulus * mThickness / (mDiameter * mSafetyFactor);
            mActualStress = mLoadPerLength * UMCLib.Pow(mHeight * 12M, 2M) * 0.25M * mDiameter / mMom;
            mCriticalYield = mYieldStrength / mSafetyFactor;

            string sPassFail = "Pass";

            if (mActualStress >= mCriticalBuckling || mActualStress >= mCriticalYield || mCriticalVelocity < mWindVP)
            {
                return "Fail";
            }

            mCriticalVelocity = UMCLib.Round(mCriticalVelocity, 4);
            mActualStress = UMCLib.Round(mActualStress, 4);
            mCriticalBuckling = UMCLib.Round(mCriticalBuckling, 4);
            mCriticalYield = UMCLib.Round(mCriticalYield, 4);
            mDeflection = UMCLib.Round(mDeflection, 4);

            return sPassFail;
        }

        #endregion Stack

        #region Underground
        /// <summary>
        /// Calculates the underground.
        /// </summary>
        /// <param name="sLoadType">Type of the s load.</param>
        /// <param name="sMaterial">The s material.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="bSpiral">if set to <c>true</c> [b spiral].</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mDistributedLoad">The m distributed load.</param>
        /// <param name="mVehicleLoad">The m vehicle load.</param>
        /// <param name="mContactArea">The m contact area.</param>
        /// <param name="mSoilDensity">The m soil density.</param>
        /// <param name="mDepth">The m depth.</param>
        /// <param name="mSoilModulus">The m soil modulus.</param>
        /// <param name="mSoilLoad">The m soil load.</param>
        /// <param name="mExternalLoad">The m external load.</param>
        /// <param name="mTotalLoad">The m total load.</param>
        /// <param name="mDeflection">The m deflection.</param>
        /// <param name="mMaxDepth">The m maximum depth.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Invalid combination of material, gauge, and duct construction.</exception>
        public static string CalcUnderground(string sLoadType, string sMaterial, int nGauge, bool bSpiral, decimal mDiameter, decimal mDistributedLoad, decimal mVehicleLoad, decimal mContactArea, decimal mSoilDensity, decimal mDepth, decimal mSoilModulus,
            out decimal mSoilLoad, out decimal mExternalLoad, out decimal mTotalLoad, out decimal mDeflection, out decimal mMaxDepth)
        {
            if (!ValidGauge(sMaterial, nGauge, bSpiral))
            {
                throw new Exception("Invalid combination of material, gauge, and duct construction.");
            }

            decimal mModulus = 0M;
            decimal mThickness = 0M;
            decimal mUnitWeight = 0M;
            decimal mYieldStrength = 0M;
            decimal mDensity = 0M;

            MetalProperties(sMaterial, nGauge, out mModulus, out mThickness, out mUnitWeight, out mYieldStrength, out mDensity);

            if (mSoilModulus <= 0M)
            {
                mSoilModulus = 200M;
            }

            switch (sLoadType)
            {
                case "Longitudinal":
                    mExternalLoad = mDistributedLoad / 12M / mDiameter;
                    break;
                case "Vehicle":
                    mExternalLoad = mVehicleLoad / mContactArea;
                    mDistributedLoad = mExternalLoad * mDiameter * 12M;
                    break;
                default:
                    mExternalLoad = 0M;
                    break;
            }

            mSoilLoad = mSoilDensity * mDepth / 144M;
            mTotalLoad = mSoilLoad + mExternalLoad;

            decimal mPipeStiffness = 0.558M * mModulus * UMCLib.Pow(mThickness / (mDiameter / 2), 3);
            decimal mCriticalPressure = 1.15M * UMCLib.Sqrt(2M * mModulus * mSoilModulus * UMCLib.Pow(mThickness, 3) / (mDiameter * (1M - 0.09M)));
            mDeflection = 0.67M * mTotalLoad / (mPipeStiffness + 0.41M * mSoilModulus) * 100M;
            mTotalLoad = mTotalLoad * 144M;

            if (mDiameter >= 4M && mDiameter <= 9.5M)
            {
                mMaxDepth = 12M * (400M - mDistributedLoad) / (mSoilDensity * mDiameter);
            }
            else if (mDiameter > 9.5M && mDiameter <= 13.5M)
            {
                mMaxDepth = 12M * (600M - mDistributedLoad) / (mSoilDensity * mDiameter);
            }
            else if (mDiameter > 13.5M && mDiameter <= 60M)
            {
                mMaxDepth = 12M * (1800M - mDistributedLoad) / (mSoilDensity * mDiameter);
            }
            else
            {
                mMaxDepth = 0M;
            }

            string sPassFail = "Fail";

            if (mDepth <= mMaxDepth && mDeflection < 10M)
            {
                sPassFail = "Pass";
            }

            mSoilLoad = UMCLib.Round(mSoilLoad, 4);
            mExternalLoad = UMCLib.Round(mExternalLoad, 4);
            mTotalLoad = UMCLib.Round(mTotalLoad, 4);
            mMaxDepth = UMCLib.Round(mMaxDepth, 4);
            mDeflection = UMCLib.Round(mDeflection, 4);

            return sPassFail;
        }

        #endregion Underground

        #region Thermal
        /// <summary>
        /// Calculates the thermal data.
        /// </summary>
        /// <param name="mInsulThickness">The m insul thickness.</param>
        /// <param name="nAmbientWind">The n ambient wind.</param>
        /// <param name="mInnerDiameter">The m inner diameter.</param>
        /// <param name="mOutsideRelHumidity">The m outside relative humidity.</param>
        /// <param name="mVolumeFlow">The m volume flow.</param>
        /// <param name="mInsideTemp">The m inside temporary.</param>
        /// <param name="mOutsideTemp">The m outside temporary.</param>
        /// <param name="mLength">Length of the m.</param>
        /// <param name="mBtuPerHour">The m btu per hour.</param>
        /// <param name="mSkinTemp">The m skin temporary.</param>
        /// <param name="mExitTemp">The m exit temporary.</param>
        /// <param name="mDewPoint">The m dew point.</param>
        /// <param name="sCondensation">The s condensation.</param>
        /// <param name="mThermConductivity">The m therm conductivity.</param>
        /// <param name="mDensity">The m density.</param>
        /// <exception cref="System.Exception">Unrecogined ambient wind speed. Valid speeds are 0, 15</exception>
        public static void CalcThermalData(decimal mInsulThickness, int nAmbientWind, decimal mInnerDiameter, decimal mOutsideRelHumidity, decimal mVolumeFlow, decimal mInsideTemp, decimal mOutsideTemp, decimal mLength,
            out decimal mBtuPerHour, out decimal mSkinTemp, out decimal mExitTemp, out decimal mDewPoint, out string sCondensation, out decimal mThermConductivity, out decimal mDensity)
        {
            decimal mHo = 0M;

            switch (nAmbientWind)
            {
                case 0:
                    mHo = 1.65M;
                    break;
                case 15:
                    mHo = 6M;
                    break;
                default:
                    throw new Exception("Unrecogined ambient wind speed. Valid speeds are 0, 15");
            }

            decimal mTempAvg = (mInsideTemp + mOutsideTemp) / 2M;
            mThermConductivity = 0.1996M + 0.0007561M * mTempAvg + 0.00000041681M * UMCLib.Pow(mTempAvg, 2M) + 0.0000000028051M * UMCLib.Pow(mTempAvg, 3M);
            if (mThermConductivity < 0.27M)
            {
                mThermConductivity = 0.27M;
            }
            if (mInsulThickness == 0M)
            {
                mThermConductivity = 350M;
            }

            decimal mAo = 2M * 3.141592M * mLength;
            mDensity = 0.075M * (530M / (mInsideTemp + 460M));
            decimal mMassflow = 60M * mDensity * mVolumeFlow;
            decimal mCp = 0.24M;

            decimal a1 = -1.044M * 10000M;
            decimal a2 = -11.29M;
            decimal a3 = -2.702M * 0.01M;
            decimal a4 = 1.289M * 0.00001M;
            decimal a5 = -2.478M * 0.000000001M;
            decimal a6 = 6.546M;
            decimal c1 = 100.45M;
            decimal c2 = 33.193M;
            decimal c3 = 2.319M;
            decimal c4 = 0.17074M;
            decimal c5 = 1.2063M;

            decimal mAbsTempOut = mOutsideTemp + 460M;
            decimal mSatPartPressOut = UMCLib.Exp(a1 / mAbsTempOut + a2 + a3 * mAbsTempOut + a4 * UMCLib.Pow(mAbsTempOut, 2M) + a5 * UMCLib.Pow(mAbsTempOut, 3M) + a6 * UMCLib.Log(mAbsTempOut));
            decimal mPartPressOut = mOutsideRelHumidity * mSatPartPressOut / 100M;
            decimal mAlfa = UMCLib.Log(mPartPressOut);
            mDewPoint = c1 + c2 * mAlfa + c3 * UMCLib.Pow(mAlfa, 2M) + c4 * UMCLib.Pow(mAlfa, 3M) + c5 * UMCLib.Pow(mPartPressOut, 0.1984M);

            decimal mLnRatio = UMCLib.Log(((mInnerDiameter + (2M * mInsulThickness)) / 2M) / (mInnerDiameter / 2M));
            decimal mSkinNumerator = mThermConductivity * (mInsideTemp - mOutsideTemp);
            decimal mSkinDenominator = mHo * (mInnerDiameter + (2M * mInsulThickness)) * mLnRatio + mThermConductivity;
            mSkinTemp = mSkinNumerator / mSkinDenominator + mOutsideTemp;

            decimal mFirstPart = 2M * 3.141592M * ((mInnerDiameter + (2M * mInsulThickness)) / 2M) * (mLength / 12M);
            decimal mSecondPart = 1M / ((((mInnerDiameter + (2M * mInsulThickness)) / 2M) / mThermConductivity) * mLnRatio + (1M / mHo));
            mBtuPerHour = mFirstPart * mSecondPart * (mInsideTemp - mOutsideTemp);
            mExitTemp = (mInsideTemp - mOutsideTemp) * UMCLib.Exp(-1M * mAo * mSecondPart / (mMassflow * mCp)) + mOutsideTemp;

            sCondensation = String.Empty;
            if (mSkinTemp <= mDewPoint + 1M)
            {
                sCondensation = "Condensation Likely";
            }
            else if (mSkinTemp > mDewPoint)
            {
                sCondensation = "Condensation Not Likely";
            }

            mBtuPerHour = UMCLib.Round(mBtuPerHour, 0);
            mSkinTemp = UMCLib.Round(mSkinTemp, 1);
            mExitTemp = UMCLib.Round(mExitTemp, 1);
            mDewPoint = UMCLib.Round(mDewPoint, 1);
            mThermConductivity = UMCLib.Round(mThermConductivity, 4);
            mDensity = UMCLib.Round(mDensity, 4);
        }
        #endregion Thermal

        #region CalcOvalRect
        /// <summary>
        /// Calculates the oval rect.
        /// </summary>
        /// <param name="sDuctType">Type of the s duct.</param>
        /// <param name="sCalcType">Type of the s calculate.</param>
        /// <param name="sPressureClass">The s pressure class.</param>
        /// <param name="mMinor">The m minor.</param>
        /// <param name="mMajor">The m major.</param>
        /// <param name="nGauge">The n gauge.</param>
        /// <param name="nSpacing">The n spacing.</param>
        /// <param name="sApplication">The s application.</param>
        /// <param name="sMinorReinforcement">The s minor reinforcement.</param>
        /// <param name="sMajorReinforcement">The s major reinforcement.</param>
        /// <param name="nCalcGauge">The n calculate gauge.</param>
        public static void CalcOvalRect(string sDuctType, string sCalcType, string sPressureClass, decimal mMinor, decimal mMajor, int nGauge, int nSpacing, string sApplication,
            out string sMinorReinforcement, out string sMajorReinforcement, out int nCalcGauge)
        {
            string sReinNoTieRods, sReinWithTieRods, sCalcSpacing;
            sMinorReinforcement = String.Empty;
            sMajorReinforcement = String.Empty;
            nCalcGauge = 0;

            switch (sCalcType)
            {
                case "Gauge":
                    CalcGaugeGivenSpacing(mMajor, nSpacing, sPressureClass, out sReinNoTieRods, out sReinWithTieRods, out nCalcGauge, out sCalcSpacing);
                    if (sReinNoTieRods == "NR")
                    {
                        sMajorReinforcement = "NR";
                    }
                    else if (nCalcGauge == 99)
                    {
                        sMajorReinforcement = "ND @ " + sCalcSpacing + " ft";
                    }
                    else
                    {
                        sMajorReinforcement = sReinNoTieRods + " @ " + sCalcSpacing + " ft";
                    }

                    CalcSpacingGivenGauge(mMinor, nCalcGauge, sPressureClass, out sReinNoTieRods, out sReinWithTieRods, out nCalcGauge, out sCalcSpacing);
                    if (sMajorReinforcement == "ND")
                    {
                        sMinorReinforcement = "ND";
                    }
                    else if (sCalcSpacing == "NR")
                    {
                        sMinorReinforcement = "NR";
                    }
                    else
                    {
                        sMinorReinforcement = sReinNoTieRods + " @ " + sCalcSpacing + " ft";
                    }
                    break;
                case "Reinforcement":
                    if (sDuctType.Contains("Oval") && sApplication == "Supply")
                    {
                        mMajor -= mMinor;
                    }

                    CalcSpacingGivenGauge(mMajor, nGauge, sPressureClass, out sReinNoTieRods, out sReinWithTieRods, out nCalcGauge, out sCalcSpacing);
                    if (sCalcSpacing == "NR")
                    {
                        sMajorReinforcement = "NR";
                    }
                    else
                    {
                        sMajorReinforcement = sReinNoTieRods + " @ " + sCalcSpacing + " ft";
                    }

                    CalcSpacingGivenGauge(mMinor, nGauge, sPressureClass, out sReinNoTieRods, out sReinWithTieRods, out nCalcGauge, out sCalcSpacing);
                    if (sCalcSpacing == "NR")
                    {
                        sMinorReinforcement = "NR";
                    }
                    else
                    {
                        sMinorReinforcement = sReinNoTieRods + " @ " + sCalcSpacing + " ft";
                    }
                    break;
            }
        }

        #endregion CalcOvalRect

        #region CalcOrificeTube
        /// <summary>
        /// Calculates the orifice tube.
        /// </summary>
        /// <param name="mPressure">The m pressure.</param>
        /// <param name="mCFM">The m CFM.</param>
        /// <param name="sPlate">The s plate.</param>
        /// <param name="mTubeDiameter">The m tube diameter.</param>
        /// <param name="mOrificeDiameter">The m orifice diameter.</param>
        /// <param name="mBetaRatio">The m beta ratio.</param>
        /// <param name="nStaticPressure">The n static pressure.</param>
        /// <param name="mLeakage">The m leakage.</param>
        /// <param name="mOpenArea">The m open area.</param>
        /// <returns></returns>
        public static decimal[,] CalcOrificeTube(decimal mPressure, decimal mCFM, string sPlate, decimal mTubeDiameter, decimal mOrificeDiameter, out decimal beta, out string sOpenArea)
        {
            decimal[,] tubeList = new decimal[8, 4];
            sOpenArea = "NONE";

            decimal dtub = mTubeDiameter;
            decimal dorf = mOrificeDiameter;
            decimal[] p = { 0.1M, 0.2M, 0.5M, 1M, 2M, 5M, 10M, 0M };
            decimal rho = 0.075M;
            decimal[] Re = new decimal[8];
            decimal[] Alfa = new decimal[8];
            decimal[] Y = new decimal[8];
            decimal[] K = new decimal[8];
            decimal[] M = new decimal[8];
            decimal[] q = new decimal[8];


            beta = dorf / dtub;
            decimal Ko = 0.5992M + 0.4252M * ((0.0006M / (UMCLib.Pow(dtub, 2M) * UMCLib.Pow(beta, 2M) + 0.01M * dtub)) + UMCLib.Pow(beta, 4M) + (1.25M * UMCLib.Pow(beta, 16M)));
            decimal b = 0.00025M + 0.002325M * (beta + 1.75M * UMCLib.Pow(beta, 4M) + (10M * UMCLib.Pow(beta, 12M)) + (2M * dtub * UMCLib.Pow(beta, 16M)));

            for (int i = 0; i < 7; i++)
            {
                Re[i] = 1363000M * dorf * UMCLib.Sqrt(rho * p[i] / (1M - UMCLib.Pow(beta, 4M)));
                Alfa[i] = p[i] / (p[i] + 13.63M * 29.71M);
                Y[i] = 1M - (0.41M + 0.35M * UMCLib.Pow(beta, 4M)) * Alfa[i] / 1.41M;
                K[i] = Ko + b * (1000M / UMCLib.Sqrt(beta * Re[i]));
                M[i] = 0.7578M * UMCLib.Pow(dorf, 2M) * Y[i] * K[i] * 0.9998M * UMCLib.Sqrt(rho * p[i] * 62.3058M);
                q[i] = M[i] / rho;
            }

            q[7] = mCFM;

            decimal area = RoundCrossSectionalArea(dtub);
            p[7] = CalcOrfPressDropAtVolFlow(q, p, mCFM);

            decimal[] vel = new decimal[8];
            decimal[] perfPressloss = new decimal[8];
            decimal pressuredrop = 0;
            decimal openarea = 0;
            if(sPlate == "1")
            {
                openarea = 23M;
                sOpenArea = openarea.ToString("N1");
            }

            for (int i = 0; i < 8; i++)
            {
                vel[i] = DuctVelocityCalc(q[i], area);
                pressuredrop = PerfPressDrop(vel[i], openarea);
                perfPressloss[i] = Math.Abs(pressuredrop);
            }

            bool bTargetPrinted = false;

            for (int i = 0; i < 8; i++)
            {
                if (q[7] <= q[i] && !bTargetPrinted)
                {
                    tubeList[i, 0] = UMCLib.Round(q[7], 1);
                    tubeList[i, 1] = UMCLib.Round(p[7], 1);
                    tubeList[i, 2] = UMCLib.Round(perfPressloss[7], 1);
                    tubeList[i, 3] = UMCLib.Round(mPressure + p[7] + perfPressloss[7], 1);
                    bTargetPrinted = true;
                }
                else if (!bTargetPrinted)
                {
                    tubeList[i, 0] = UMCLib.Round(q[i], 1);
                    tubeList[i, 1] = UMCLib.Round(p[i], 1);
                    tubeList[i, 2] = UMCLib.Round(perfPressloss[i], 1);
                    tubeList[i, 3] = UMCLib.Round(mPressure + p[i] + perfPressloss[i], 1);
                }
                else
                {
                    tubeList[i, 0] = UMCLib.Round(q[i - 1], 1);
                    tubeList[i, 1] = UMCLib.Round(p[i - 1], 1);
                    tubeList[i, 2] = UMCLib.Round(perfPressloss[i - 1], 1);
                    tubeList[i, 3] = UMCLib.Round(mPressure + p[i - 1] + perfPressloss[i - 1], 1);
                }
            }

            return tubeList;
        }
        #endregion CalcOrificeTube

        #region CalcDuctDFuser
        /// <summary>
        /// Calculates the duct d fuser.
        /// </summary>
        /// <param name="nVelocity">The n velocity.</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="bConstantDiameter">if set to <c>true</c> [b constant diameter].</param>
        /// <param name="mCFM">The m CFM.</param>
        /// <param name="mLength">Length of the m.</param>
        /// <param name="ductList">The duct list.</param>
        /// <param name="orificeList">The orifice list.</param>
        /// <param name="reducerList">The reducer list.</param>
        /// <returns></returns>
        public static string CalcDuctDFuser(int nVelocity, decimal mDiameter, bool bConstantDiameter, decimal mCFM, decimal mLength,
            out List<string> ductList, out List<string> orificeList, out int nEnteringVelocity, out decimal mPressureDrop)
        {
            string sError = String.Empty;

            decimal q = mCFM;
            int d = 0;

            if (mDiameter <= 0)
            {
                d = (int)UMCLib.Round(UMCLib.Pow((q / nVelocity) / 3.14159M, .5M) * 24M, 0);
            }
            else
            {
                d = (int)UMCLib.Round(mDiameter, 0);
            }

            if (d > 38 && d % 2 != 0)
            {
                d -= 1;
            }

            int tl = (int)UMCLib.Round(mLength, 0);
            int l = (int)UMCLib.Round((q / 50M) / ((d / 12M) * 3.1415927M), 0);

            if (tl < l)
            {
                sError = "Total Length Too Small, Designing with Minimum";
                tl = l;
            }

            int n = (int)UMCLib.Round(tl / 4.5M, 0);
            if (n < 4)
            {
                n = 4;
            }
            else if (n > 10)
            {
                n = 10;
            }

            int sl = tl / n;
            decimal dq = q / n;
            tl = 0;

            decimal area = (3.14159M * UMCLib.Pow(d, 2M)) / 576M;
            decimal fpm = q / area;
            decimal tp = UMCLib.Pow(fpm / 4005M, 2M) + 0.1M;

            decimal d2 = 0;
            decimal[] da = new decimal[n];
            decimal[] dd = new decimal[n];
            string[] rd = new string[n];

            for (int i = 1; i < n; i++)
            {
                tl = tl + sl;

                if (!bConstantDiameter && n > 4 && (i == 4 || i == 7 || i == 10))
                {
                    d = (int)UMCLib.Round(d2, 0);
                    if (d > 38 && d % 2 != 0)
                    {
                        d -= 1;
                    }
                }

                decimal qup = q - ((i - 1) * dq);
                decimal qdn = q - (i * dq);
                decimal a = UMCLib.Pow(d / 24M, 2M) * 3.14159M;
                decimal b = (qup - qdn) / qup;
                decimal m = (qdn * 0.075M) / 60M;
                decimal a2 = a - (b * a);
                d2 = UMCLib.Pow(a2 / 3.14159M, 0.5M) * 24M;
                decimal x = -1M;

                while (x != 0M)
                {
                    decimal bt = d2 / d;
                    a2 = UMCLib.Pow(d2 / 24M, 2M) * 3.14159M;
                    decimal v2 = qdn / a2;
                    decimal r = (48M * m) / (3.14159M * d2 * .0000121M);
                    decimal lambda = 1000M / UMCLib.Pow(bt * r, 0.5M);
                    decimal e = 1M / UMCLib.Pow(1M - UMCLib.Pow(bt, 4M), 0.5M);
                    decimal ko = (0.6014M - (0.01352M * UMCLib.Pow(d, -0.25M))) + (0.376M + (0.07257001M * UMCLib.Pow(d, -0.25M))) * ((0.00025M / (((d ^ 2) * UMCLib.Pow(bt, 2M)) + (0.0025M * d))) + UMCLib.Pow(bt, 4M) + (1.5M * UMCLib.Pow(bt, 16M)));
                    decimal b1 = (0.0002M + (0.0011M / d)) + ((0.0038M + (0.0004M / d)) * (UMCLib.Pow(bt, 2M) + ((16.5M + (5M * d)) * UMCLib.Pow(bt, 16M))));
                    decimal k = ko + (b1 * lambda);
                    decimal kc = k / e;
                    decimal p1 = 14.696M;
                    decimal v1 = qup / a;
                    decimal vp1 = UMCLib.Pow(v1 / 4005M, 2M);
                    decimal v3 = qdn / a;
                    decimal vp3 = UMCLib.Pow(v3 / 4005M, 2M);
                    decimal mpd = vp1 - vp3;
                    decimal pd = mpd;
                    decimal p2 = 14.696M - (pd / 27.7M);
                    decimal xz = 1M - (p2 / p1);
                    decimal gamma = 1.41M;
                    decimal y = 1M - ((0.41M + (0.35M * UMCLib.Pow(bt, 4M))) * (xz / gamma));
                    decimal dp = UMCLib.Pow(m / (0.099702M * e * y * UMCLib.Pow(d2, 2M) * kc), 2M) / 0.075M;
                    decimal kb = 1.0005M - (0.1806M * bt) - (0.7174M * UMCLib.Pow(bt, 2M));
                    decimal opd = kb * dp;
                    x = opd - mpd;

                    if (x != 0)
                    {
                        if (x < 0)
                        {
                            x = Math.Abs(x);

                            if (x > 0.01M)
                            {
                                d2 -= 0.05M;
                            }
                            else
                            {
                                x = 0M;
                            }
                        }
                        else
                        {
                            if (x > 0.01M)
                            {
                                d2 += 0.05M;
                            }
                            else
                            {
                                x = 0M;
                            }
                        }
                    }
                }

                da[i - 1] = d;

                if (!bConstantDiameter && n >= 4 && (i == 3 || i == 6 || i == 9))
                {
                    d = (int)UMCLib.Round(d2, 0);
                    if (d > 38 && d % 2 != 0)
                    {
                        d -= 1;
                    }

                    dd[i - 1] = d;
                }
                else
                {
                    dd[i - 1] = d2;
                }

                if (i == 3 || i == 6 || i == 9)
                {
                    rd[i - 1] = " R";
                }
            }

            if (!bConstantDiameter)
            {
                switch (n)
                {
                    case 4:
                        da[n - 1] = dd[n - 2];
                        break;
                    case 5:
                        da[n - 1] = da[n - 2];
                        break;
                }
            }
            else
            {
                for (int i = 0; i < rd.Length; i++)
                {
                    rd[i] = String.Empty;
                }
            }

            tl = tl + sl;

            ductList = new List<string>();

            for (int i = 0; i < da.Length; i++)
            {
                decimal mValue = da[i];

                if (mValue == 0 && i == da.Length - 1)
                {
                    mValue = da[i - 1];
                }

                ductList.Add(UMCLib.Round(mValue, 0).ToString());
            }

            orificeList = new List<string>();

            for (int i = 0; i < dd.Length; i++)
            {
                decimal mValue = dd[i];
                string sReducer = rd[i];

                if (mValue == 0M)
                {
                    orificeList.Add("End Cap");
                }
                else if (!String.IsNullOrEmpty(sReducer))
                {
                    orificeList.Add(UMCLib.Round(mValue, 0).ToString() + ".00" + sReducer);
                }
                else
                {
                    orificeList.Add(UMCLib.Round(mValue, 2).ToString());
                }
            }

            nEnteringVelocity = (int)UMCLib.Round(fpm, 0);
            mPressureDrop = UMCLib.Round(tp, 2);

            return sError;
        }

        #endregion CalcDuctDFuser

        #region CalcFactair
        /// <summary>
        /// Calculates the factair.
        /// </summary>
        /// <param name="nModel">The n model.</param>
        /// <param name="mVelocity">The m velocity.</param>
        /// <param name="bOpen">if set to <c>true</c> [b open].</param>
        /// <param name="nDistance">The n distance.</param>
        /// <param name="mMaxVelocity">The m maximum velocity.</param>
        /// <param name="mPressureDrop">The m pressure drop.</param>
        /// <param name="octaves">The octaves.</param>
        public static void CalcFactair(int nModel, decimal mVelocity, bool bOpen, int nDistance, out decimal mMaxVelocity, out decimal mPressureDrop, out decimal[] octaves)
        {
            decimal mMaxVel25 = 0M, mMaxVel50 = 0M, mMaxVel75 = 0M, mMaxVel100 = 0M;
            decimal mPDrop25 = 0M, mPDrop50 = 0M, mPDrop75 = 0M, mPDrop100 = 0M;
            decimal mOctave1 = 0M, mOctave2 = 0M, mOctave3 = 0M, mOctave4 = 0M, mOctave5 = 0M, mOctave6 = 0M, mOctave7 = 0M, mOctave8 = 0M;

            if (bOpen)
            {
                if (nDistance == 2)
                {
                    mMaxVel25 = -16.4M + 0.9M * mVelocity + 2.1E-5M * UMCLib.Pow(mVelocity, 2);
                    mMaxVel50 = -6.4M + 0.76M * mVelocity + 4.3E-5M * UMCLib.Pow(mVelocity, 2);
                    mMaxVel75 = -5.7M + 0.78M * mVelocity + 2.1E-5M * UMCLib.Pow(mVelocity, 2);
                    mMaxVel100 = -5M + 0.78M * mVelocity + 1.64E-5M * UMCLib.Pow(mVelocity, 2);
                }
                else if (nDistance == 5)
                {
                    mMaxVel25 = 2.1M + 0.48M * mVelocity + 3.3E-5M * UMCLib.Pow(mVelocity, 2);
                    mMaxVel50 = 3.57M + 0.43M * mVelocity + 4E-5M * UMCLib.Pow(mVelocity, 2);
                    mMaxVel75 = 1.07M + 0.35M * mVelocity + 5.7E-5M * UMCLib.Pow(mVelocity, 2);
                    mMaxVel100 = 3.33M + 0.31M * mVelocity + 6.17E-5M * UMCLib.Pow(mVelocity, 2);
                }

                mPDrop25 = -0.001M + 1.30E-5M * mVelocity + 4.82E-13M * UMCLib.Pow(mVelocity, 3);
                mPDrop50 = -0.002M + 1.25E-5M * mVelocity + 4.34E-12M * UMCLib.Pow(mVelocity, 3);
                mPDrop75 = -0.002M + 1.25E-5M * mVelocity + 4.34E-12M * UMCLib.Pow(mVelocity, 3);
                mPDrop100 = -0.003M + 2.2E-5M * mVelocity + 5.91E-12M * UMCLib.Pow(mVelocity, 3);
            }
            else
            {
                switch (nDistance)
                {
                    case 1:
                        mMaxVel25 = 1.07M + 0.314M * mVelocity - 2.43E-5M * UMCLib.Pow(mVelocity, 2);
                        break;
                    case 2:
                        mMaxVel25 = 0.298M + 0.217M * mVelocity - 2.83E-5M * UMCLib.Pow(mVelocity, 2);
                        break;
                    case 3:
                        mMaxVel25 = 0.417M + 0.144M * mVelocity - 1.6E-5M * UMCLib.Pow(mVelocity, 2);
                        break;
                    case 4:
                        mMaxVel25 = 0.119M + 0.0976M * mVelocity - 7.62E-6M * UMCLib.Pow(mVelocity, 2);
                        break;
                    case 5:
                        mMaxVel25 = 0.357M + 0.0674M * mVelocity - 4.29E-6M * UMCLib.Pow(mVelocity, 2);
                        break;
                    case 6:
                        mMaxVel25 = 0.238M + 0.0433M * mVelocity + 4.76E-7M * UMCLib.Pow(mVelocity, 2);
                        break;
                    case 10:
                        mMaxVel25 = 0.119M + 0.0207M * mVelocity - 1.905E-6M * UMCLib.Pow(mVelocity, 2);
                        break;
                }

                mMaxVel50 = mMaxVel25;
                mMaxVel75 = mMaxVel25;
                mMaxVel100 = mMaxVel25;

                mPDrop25 = -0.001M + 1.58E-5M * mVelocity + 6.17E-12M * UMCLib.Pow(mVelocity, 3);
                mPDrop50 = -0.004M + 3.45E-5M * mVelocity + 1.02E-11M * UMCLib.Pow(mVelocity, 3);
                mPDrop75 = -0.0068M + 3.67E-5M * mVelocity + 1.35E-12M * UMCLib.Pow(mVelocity, 3);
                mPDrop100 = -0.0045M + 3.40E-5M * mVelocity + 8.13E-12M * UMCLib.Pow(mVelocity, 3);
            }

            mPressureDrop = 0M;
            mMaxVelocity = 0M;

            switch (nModel)
            {
                case 25:
                    mPressureDrop = UMCLib.Round(mPDrop25, 3);
                    mMaxVelocity = UMCLib.Round(mMaxVel25, 3);
                    mOctave1 = 74.8M - 74.8M * UMCLib.Exp(-1 * mVelocity / 1195M);
                    mOctave2 = 77.4M - 77.4M * UMCLib.Exp(-1 * mVelocity / 1364M);
                    mOctave3 = 72.4M - 72.4M * UMCLib.Exp(-1 * mVelocity / 1276M);
                    mOctave4 = 68.3M - 68.3M * UMCLib.Exp(-1 * mVelocity / 1041M);
                    mOctave5 = 72.9M - 72.9M * UMCLib.Exp(-1 * mVelocity / 1126M);
                    mOctave6 = 75.9M - 75.9M * UMCLib.Exp(-1 * mVelocity / 1413M);
                    mOctave7 = 78.0M - 78.0M * UMCLib.Exp(-1 * mVelocity / 1750M);
                    mOctave8 = 92.9M - 93.0M * UMCLib.Exp(-1 * mVelocity / 2714M);
                    break;
                case 50:
                    mPressureDrop = UMCLib.Round(mPDrop50, 3);
                    mMaxVelocity = UMCLib.Round(mMaxVel50, 3);
                    mOctave1 = 74.8M - 74.8M * UMCLib.Exp(-1 * mVelocity / 1195M);
                    mOctave2 = 75.6M - 75.6M * UMCLib.Exp(-1 * mVelocity / 1178M);
                    mOctave3 = 77.7M - 77.6M * UMCLib.Exp(-1 * mVelocity / 1249M);
                    mOctave4 = 75.5M - 75.4M * UMCLib.Exp(-1 * mVelocity / 1099M);
                    mOctave5 = 77.3M - 77.3M * UMCLib.Exp(-1 * mVelocity / 1146M);
                    mOctave6 = 78.0M - 78.0M * UMCLib.Exp(-1 * mVelocity / 1317M);
                    mOctave7 = 84.2M - 84.3M * UMCLib.Exp(-1 * mVelocity / 1808M);
                    mOctave8 = 104.7M - 104.8M * UMCLib.Exp(-1 * mVelocity / 2971M);
                    break;
                case 75:
                    mPressureDrop = UMCLib.Round(mPDrop75, 3);
                    mMaxVelocity = UMCLib.Round(mMaxVel75, 3);
                    mOctave1 = 74.8M - 74.8M * UMCLib.Exp(-1 * mVelocity / 1195);
                    mOctave2 = 77.4M - 77.4M * UMCLib.Exp(-1 * mVelocity / 1364);
                    mOctave3 = 72.4M - 72.4M * UMCLib.Exp(-1 * mVelocity / 1276);
                    mOctave4 = 68.3M - 68.3M * UMCLib.Exp(-1 * mVelocity / 1041);
                    mOctave5 = 72.9M - 72.9M * UMCLib.Exp(-1 * mVelocity / 1126);
                    mOctave6 = 75.9M - 75.9M * UMCLib.Exp(-1 * mVelocity / 1413);
                    mOctave7 = 78.0M - 78.0M * UMCLib.Exp(-1 * mVelocity / 1750);
                    mOctave8 = 92.9M - 93.0M * UMCLib.Exp(-1 * mVelocity / 2714);
                    break;
                case 100:
                    mPressureDrop = UMCLib.Round(mPDrop100, 3);
                    mMaxVelocity = UMCLib.Round(mMaxVel100, 3);
                    mOctave1 = 74.8M - 74.8M * UMCLib.Exp(-1 * mVelocity / 1195M);
                    mOctave2 = 75.6M - 75.6M * UMCLib.Exp(-1 * mVelocity / 1178M);
                    mOctave3 = 77.7M - 77.6M * UMCLib.Exp(-1 * mVelocity / 1249M);
                    mOctave4 = 75.5M - 75.4M * UMCLib.Exp(-1 * mVelocity / 1099M);
                    mOctave5 = 77.3M - 77.3M * UMCLib.Exp(-1 * mVelocity / 1146M);
                    mOctave6 = 78.0M - 78.0M * UMCLib.Exp(-1 * mVelocity / 1317M);
                    mOctave7 = 84.2M - 84.3M * UMCLib.Exp(-1 * mVelocity / 1808M);
                    mOctave8 = 104.7M - 104.8M * UMCLib.Exp(-1 * mVelocity / 2971M);
                    break;
            }

            if (mMaxVelocity < 0M)
            {
                mMaxVelocity = 0M;
            }

            if (mPressureDrop < 0M)
            {
                mPressureDrop = 0M;
            }

            octaves = new decimal[8];
            octaves[0] = UMCLib.Round(mOctave1, 1);
            octaves[1] = UMCLib.Round(mOctave2, 1);
            octaves[2] = UMCLib.Round(mOctave3, 1);
            octaves[3] = UMCLib.Round(mOctave4, 1);
            octaves[4] = UMCLib.Round(mOctave5, 1);
            octaves[5] = UMCLib.Round(mOctave6, 1);
            octaves[6] = UMCLib.Round(mOctave7, 1);
            octaves[7] = UMCLib.Round(mOctave8, 1);
        }

        #endregion CalcFactair

        #region CalcOffset
        /// <summary>
        /// Calculates the offset.
        /// </summary>
        /// <param name="sCalcType">Type of the s calculate.</param>
        /// <param name="sConnectionType">Type of the s connection.</param>
        /// <param name="mDiameter">The m diameter.</param>
        /// <param name="mDistance">The m distance.</param>
        /// <param name="mLength">Length of the m.</param>
        /// <param name="mCenterline1">The m centerline1.</param>
        /// <param name="mCenterline2">The m centerline2.</param>
        /// <param name="mAngle">The m angle.</param>
        /// <param name="mDuctLength">Length of the m duct.</param>
        public static void CalcOffset(string sCalcType, string sConnectionType, decimal mDiameter, decimal mLength, decimal mCenterline1, decimal mCenterline2,
            ref decimal mDistance, ref decimal mAngle, out decimal mCalcLength)
        {
            if (sCalcType == "Duct")
            {
                mAngle = 45M;
                mCalcLength = (mDistance / UMCLib.Tan(mAngle)) + (mCenterline1 * mDiameter * UMCLib.Sin(mAngle / 2M)) + (mCenterline2 * mDiameter * UMCLib.Tan(mAngle / 2M));

                for (int i = 0; i < 4500 && Math.Abs(mCalcLength - mLength) >= 0.1M; i++)
                {
                    if (mCalcLength < mLength)
                    {
                        mAngle -= 0.01M;
                    }
                    else
                    {
                        mAngle += 0.01M;
                    }

                    if (mAngle > 90M || mAngle < 0M)
                    {
                        throw new Exception("Too many iterations - invalid entry");
                    }

                    mCalcLength = (mDistance / UMCLib.Tan(mAngle)) + (mCenterline1 * mDiameter * UMCLib.Sin(mAngle / 2M)) + (mCenterline2 * mDiameter * UMCLib.Tan(mAngle / 2M));
                }

                decimal mAdj;

                switch (sConnectionType)
                {
                    case "ring":
                        mAdj = -2M * 2.25M;
                        break;
                    case "vanstone":
                        mAdj = -2M * 1.5M;
                        break;
                    default:
                        mAdj = 0M;
                        break;
                }

                mCalcLength = (mDistance / UMCLib.Sin(mAngle)) - (mCenterline1 * mDiameter * UMCLib.Tan(mAngle / 2M) + mCenterline2 * mDiameter * UMCLib.Tan(mAngle / 2M));

                if (mDiameter < 22M)
                {
                    mCalcLength += mAdj;
                }
                else
                {
                    mCalcLength += 4 + mAdj;
                }
            }
            else
            {
                if (mDiameter < 22)
                {
                    mDistance = 2M * mCenterline1 * mDiameter * (1M - UMCLib.Cos(mAngle)) + 4M * UMCLib.Sin(mAngle);
                    mCalcLength = 2M * mCenterline1 * mDiameter * UMCLib.Sin(mAngle) + 4M * UMCLib.Cos(mAngle);
                }
                else
                {
                    mDistance = 2M * mCenterline1 * mDiameter * (1M - UMCLib.Cos(mAngle));
                    mCalcLength = 2M * mCenterline1 * mDiameter * UMCLib.Sin(mAngle) - 4M;
                }
            }

            mDistance = UMCLib.Round(mDistance, 2);
            mAngle = UMCLib.Round(mAngle, 0);
            mCalcLength = UMCLib.Round(mCalcLength, 2);
        }

        #endregion CalcOffset

        #region CalcAcoustical
        /// <summary>
        /// Calculates the acoustical.
        /// </summary>
        /// <param name="sCalcType">Type of the s calculate.</param>
        /// <param name="nLevels">Number of levels.</parm>
        /// <param name="InputLevels">The input levels.</param>
        /// <param name="dLengthBefore">The d length before.</param>
        /// <param name="dLengthAfter">The d length after.</param>
        /// <param name="dDistance">The d distance.</param>
        /// <param name="OutputLevels">The output levels.</param>
        /// <param name="dOverall">The d overall.</param>
        public static void CalcAcoustical(string sCalcType, int nColumns, int[,] InputLevels, double dLengthBefore, double dLengthAfter, double dDistance, 
            out int[,] OutputLevels, out double dOverall)
        {
            int[] dbAdj = { 26, 16, 9, 3, 0, 1, 1, -1 };
            double dCalcLogQty = 0;
            OutputLevels = null;
            dOverall = 0;
            int nRows =  InputLevels.GetLength(1);

            if (sCalcType == "AWeighting")
            {
                OutputLevels = new int[1, nRows];

                for (int i = 0; i < nRows; i++)
                {
                    OutputLevels[0, i] = InputLevels[0, i] - dbAdj[i];
                    dCalcLogQty += Math.Pow(10, 0.1 * OutputLevels[0, i]);
                }

                dOverall = 10D * Math.Log10(dCalcLogQty);
            }
            else if (sCalcType == "Addition")
            {
                OutputLevels = new int[2, nRows];

                double[] octaves = new double[nRows];
                double dCalcLogQty2 = 0;

                for (int i = 0; i < nRows; i++)
                {
                    //for (int j = 0; j < InputLevels.GetLength(0); j++)
                    for (int j = 0; j < nColumns; j++)
                    {
                        dCalcLogQty += Math.Pow(10, 0.1 * InputLevels[j, i]);
                    }

                    octaves[i] = 10D * Math.Log10(dCalcLogQty);
                    dCalcLogQty = 0;

                    OutputLevels[0, i] = (int)Math.Round(octaves[i], 0);
                    OutputLevels[1, i] = OutputLevels[0, i] - dbAdj[i];
                    dCalcLogQty2 += Math.Pow(10, 0.1 * OutputLevels[1, i]);
                }

                dOverall = 10D * Math.Log10(dCalcLogQty2);
            }
            else if (sCalcType == "Discharge")
            {
                OutputLevels = new int[1, nRows];

                double[] SoundInSilencer = new double[nRows];
                double[] SoundExitSilencerNoGNL = new double[nRows];
                double[] SoundExitSilencerWithGNL = new double[nRows];
                double[] SoundExit = new double[nRows];

                for (int i = 0; i < nRows; i++)
                {
                    SoundInSilencer[i] = InputLevels[0, i] - InputLevels[1, i] * dLengthBefore;
                    SoundExitSilencerNoGNL[i] = SoundInSilencer[i] - InputLevels[2, i];
                    if (SoundExitSilencerNoGNL[i] < 0)
                    {
                        SoundExitSilencerNoGNL[i] = 0;
                    }
                    SoundExitSilencerWithGNL[i] = 10 * Math.Log10(Math.Pow(10, (.1 * SoundExitSilencerNoGNL[i])) + Math.Pow(10, (.1 * InputLevels[3, i])));
                    SoundExit[i] = SoundExitSilencerWithGNL[i] - InputLevels[4, i] * dLengthAfter;
                    if (SoundExit[i] < 0)
                    {
                        SoundExit[i] = 0;
                    }
                    OutputLevels[0, i] = (int)Math.Round(SoundExit[i] - 20 * Math.Log10(dDistance) + 2.3);
                    if (OutputLevels[0, i] < 0)
                    {
                        OutputLevels[0, i] = 0;
                    }
                    int nDBAOctave = OutputLevels[0, i] - dbAdj[i];
                    dCalcLogQty += Math.Pow(10, 0.1 * nDBAOctave);
                }

                dOverall = 10D * Math.Log10(dCalcLogQty);
            }

            if (dOverall < 0)
            {
                dOverall = 0;
            }
            else
            {
                dOverall = UMCLib.Round(dOverall, 2);
            }
        }

        #endregion CalcAcoustical

        private static void CalcSpacingGivenGauge(decimal mPanelWidth, int nGauge, string sPressureClass,
            out string sReinNoTieRods, out string sReinWithTieRods, out int nCalcGauge, out string sCalcSpacing)
        {
            string[] foundData = FindRectRecord(mPanelWidth, sPressureClass);
            string sRawGauge = String.Empty;
            int i = 0;
            sCalcSpacing = String.Empty;
            nCalcGauge = 0;

            for (i = 3; i < 12; i++)
            {
                sRawGauge = foundData[i].Trim('"');
                nCalcGauge = ConvertRawGauge(sRawGauge);

                if (nCalcGauge == nGauge)
                {
                    break;
                }
                else
                {
                    nCalcGauge = 0;
                }
            }

            switch (i)
            {
                case 3:
                    sCalcSpacing = "NR";
                    break;
                case 4:
                    sCalcSpacing = "10";
                    break;
                case 5:
                    sCalcSpacing = "8";
                    break;
                case 6:
                    sCalcSpacing = "6";
                    break;
                case 7:
                    sCalcSpacing = "5";
                    break;
                case 8:
                    sCalcSpacing = "4";
                    break;
                case 9:
                    sCalcSpacing = "3";
                    break;
                case 10:
                    sCalcSpacing = "2.5";
                    break;
                case 11:
                    sCalcSpacing = "2";
                    break;
                default:
                    sRawGauge = String.Empty;
                    break;
            }

            if (!String.IsNullOrEmpty(sRawGauge))
            {
                CalcReinforcementTypeOptions(sRawGauge, out sReinNoTieRods, out sReinWithTieRods);
                return;
            }

            sRawGauge = foundData[3];
            int nLowGauge = ConvertRawGauge(sRawGauge);

            if (nGauge < nLowGauge)
            {
                sCalcSpacing = "NR";
                nCalcGauge = nLowGauge;
                CalcReinforcementTypeOptions(sRawGauge, out sReinNoTieRods, out sReinWithTieRods);
            }
            else
            {
                for (i = 4; i < 11; i++)
                {
                    sRawGauge = foundData[i].Trim('"');
                    nCalcGauge = ConvertRawGauge(sRawGauge);

                    if (nGauge < nCalcGauge)
                    {
                        break;
                    }
                    else
                    {
                        nCalcGauge = 0;
                    }
                }

                switch (i)
                {
                    case 4:
                        sCalcSpacing = "10";
                        break;
                    case 5:
                        sCalcSpacing = "8";
                        break;
                    case 6:
                        sCalcSpacing = "6";
                        break;
                    case 7:
                        sCalcSpacing = "5";
                        break;
                    case 8:
                        sCalcSpacing = "4";
                        break;
                    case 9:
                        sCalcSpacing = "3";
                        break;
                    case 10:
                        sCalcSpacing = "2.5";
                        break;
                    case 11:
                        sCalcSpacing = "2";
                        break;
                }

                CalcReinforcementTypeOptions(sRawGauge, out sReinNoTieRods, out sReinWithTieRods);
            }
        }

        private static int ConvertRawGauge(string sRawGauge)
        {
            sRawGauge = sRawGauge.Trim('"');

            if (sRawGauge == "ND")
            {
                return 0;
            }

            string sGauge = sRawGauge;

            while (Char.IsLetter(sGauge[0]))
            {
                sGauge = sGauge.Substring(1);
            }

            sGauge = sGauge.Substring(0, 2);
            return UMCLib.ConvertToInt32(sGauge);
        }

        private static void CalcGaugeGivenSpacing(decimal mPanelWidth, int nSpacing, string sPressureClass,
            out string sReinNoTieRods, out string sReinWithTieRods, out int nCalcGauge, out string sCalcSpacing)
        {
            string[] foundData = FindRectRecord(mPanelWidth, sPressureClass);

            if (foundData == null)
            {
                sReinNoTieRods = String.Empty;
                sReinWithTieRods = String.Empty;
                nCalcGauge = 0;
                sCalcSpacing = String.Empty;
                return;
            }

            string sRawGauge = foundData[3 + nSpacing].Trim('"');

            if (sRawGauge == "ND")
            {
                sCalcSpacing = "Need to increase gauge";
                nCalcGauge = 99;
            }
            else
            {
                nCalcGauge = ConvertRawGauge(sRawGauge);

                switch (nSpacing)
                {
                    case 0:
                        sCalcSpacing = "NR";
                        break;
                    case 1:
                        sCalcSpacing = "10";
                        break;
                    case 2:
                        sCalcSpacing = "8";
                        break;
                    case 3:
                        sCalcSpacing = "6";
                        break;
                    case 4:
                        sCalcSpacing = "5";
                        break;
                    case 5:
                        sCalcSpacing = "4";
                        break;
                    case 6:
                        sCalcSpacing = "3";
                        break;
                    case 7:
                        sCalcSpacing = "2.5";
                        break;
                    default:
                        sCalcSpacing = "2";
                        break;
                }
            }

            CalcReinforcementTypeOptions(sRawGauge, out sReinNoTieRods, out sReinWithTieRods);
        }

        private static string[] FindRectRecord(decimal mPanelWidth, string sPressureClass)
        {
            List<string[]> rectData = ReadRectData();
            string[] foundData = null;
            char[] badChars = { '"', ' ' };

            foreach (string[] data in rectData)
            {
                if (data[0].Trim(badChars) == sPressureClass)
                {
                    int nLow = UMCLib.ConvertToInt32(data[1]);
                    int nHigh = UMCLib.ConvertToInt32(data[2]);

                    if (mPanelWidth >= nLow && mPanelWidth <= nHigh)
                    {
                        foundData = data;
                        break;
                    }
                }
            }
            return foundData;
        }

        private static void CalcReinforcementTypeOptions(string sRawGauge, out string sReinNoIeRods, out string sReinWithTieRods)
        {
            if (sRawGauge[1] == 't')
            {
                sReinNoIeRods = sRawGauge.Substring(0, 2);
                sReinWithTieRods = sRawGauge[sRawGauge.Length - 1].ToString();
            }
            else if (sRawGauge[0] == sRawGauge[sRawGauge.Length - 1])
            {
                sReinNoIeRods = sRawGauge;
                sReinWithTieRods = String.Empty;
            }
            else if (sRawGauge == "ND")
            {
                sReinNoIeRods = sRawGauge;
                sReinWithTieRods = sRawGauge;
            }
            else
            {
                sReinNoIeRods = sRawGauge[0].ToString();
                sReinWithTieRods = sRawGauge[sRawGauge.Length - 1].ToString();
            }
        }

        private static List<string[]> ReadRectData()
        {
            if (_RectData == null)
            {
                _RectData = new List<string[]>();

                foreach (string sLine in File.ReadAllLines(@"C:\umcdata\RECTNGLR.DAT"))
                {
                    _RectData.Add(sLine.Split(','));
                }
            }

            return _RectData;
        }

        private static decimal MomentOfInertia(bool bSpiral, decimal mDiameter, decimal mThickness)
        {
            decimal mMomFactor = 0.04M;
            if (bSpiral)
            {
                mMomFactor = 0.02M;
            }

            decimal mMom = mMomFactor * (UMCLib.Pow(mDiameter, 4M) - UMCLib.Pow(mDiameter - 2 * mThickness, 4M));
            return mMom;
        }

        private static void StackDeflection(decimal mDiameter, decimal mLoadPerLength, decimal mHeight, decimal mModulus, decimal mMom,
             out decimal mDeflection, out decimal mMinumumThickness)
        {
            mDeflection = mLoadPerLength * UMCLib.Pow(mHeight * 12M, 4M) / (8M * mModulus * mMom);
            mMinumumThickness = 0.5M * (mDiameter - UMCLib.Pow(UMCLib.Pow(mDiameter, 4M) - (720M * mLoadPerLength * UMCLib.Pow(mHeight * 12M, 3M)) / (0.16M * mModulus), 0.25M));
        }

        private static void SupportDeflection(decimal mDiameter, decimal mLoadPerLength, decimal mLength, decimal mModulus, decimal mMom, decimal mSafetyFactor,
            out decimal mDeflection, out decimal mAllowedDeflection, out decimal mMaxLength)
        {
            mDeflection = 5M * mLoadPerLength * UMCLib.Pow(mLength * 12M, 4M) / (384M * mModulus * mMom);
            mAllowedDeflection = mLength * 12M / (240M * mSafetyFactor);
            mMaxLength = (1M / 12M) * UMCLib.Pow((mModulus * mMom * 384M) / (240M * mSafetyFactor * 5M * mLoadPerLength), 1M / 3M);

            mDeflection = UMCLib.Round(mDeflection, 4);
            mAllowedDeflection = UMCLib.Round(mAllowedDeflection, 4);
            mMaxLength = UMCLib.Round(mMaxLength, 4);
        }

        private static decimal ClassThickness(string sClass)
        {
            switch (sClass)
            {
                case "Class B":
                case "Class 2":
                    return 0.006M;
                case "Class C":
                case "Class 3":
                    return 0.014M;
                case "Class D":
                case "Class 4":
                    return 0.024M;
                default :
                    return 0M;
            }
        }

        private static List<string[]> ReadAngleData()
        {
            if (_AngleData == null)
            {
                _AngleData = new List<string[]>();

                foreach (string sLine in File.ReadAllLines(@"C:\unitedmcgill\wwwroot\bin\ANGLE.DAT"))
                {
                    _AngleData.Add(sLine.Split(','));
                }
            }

            return _AngleData;
        }

        private static string StiffenerSize(decimal mDiameter, decimal mPressure, decimal mStiffenerSpacing, decimal mModulus)
        {
            string sStiffenerSize = String.Empty;

            decimal mArea = (1.031M * UMCLib.Pow(10M, -6M)) * mPressure * mStiffenerSpacing * mDiameter * (52M + mDiameter);
            decimal mMoment = (0.397M * mPressure * mStiffenerSpacing * UMCLib.Pow(mDiameter, 3M)) / mModulus;

            if (mDiameter >= 6M && mDiameter <= 42M && mArea <= 0.2516M && mMoment <= 0.0383M)
            {
                sStiffenerSize = "UNI-RING or UNI-FLANGE";
            }
            else if (mDiameter >= 43M && mDiameter <= 60M && mArea <= 0.3173M && mMoment <= 0.0474M)
            {
                sStiffenerSize = "UNI-RING may be used";
            }
            else if (mArea <= 0.241M && mMoment < 0.0764M)
            {
                sStiffenerSize = "UNI-FLANGE may be used";
            }
            else if(mDiameter > 36M && mArea <= 0.71M && mMoment <= 0.27M)
            {
                sStiffenerSize = "2 x 2 x 3/16";
            }
            else if (mDiameter >= 3M && mDiameter <= 5M)
            {
                sStiffenerSize = "1 x 1 x 10 ga.";
            }
            else if (mDiameter >= 6M && mDiameter <= 11M)
            {
                sStiffenerSize = "1 1/4 x 1 1/4 x 1/8";
            }
            else if (mDiameter >= 12M && mDiameter <= 15M)
            {
                sStiffenerSize = "1 1/2 x 1 1/2 x 1/8";
            }
            else if (mDiameter >= 16M && mDiameter <= 24M)
            {
                sStiffenerSize = "1 1/2 x 1 1/2 x 3/16";
            }
            else if (mDiameter >= 25M && mDiameter <= 36M)
            {
                sStiffenerSize = "2 x 2 x 3/16";
            }
            else if (mDiameter > 36M)
            {
                List<string[]> angleData = ReadAngleData();
                for (int i = 0; i < angleData.Count; i++)
                {
                    string[] curData = angleData[i];
                    if(curData.Length >2)
                    {
                        if (mArea <= UMCLib.ConvertToDecimal(curData[1]) && mMoment <= UMCLib.ConvertToDecimal(curData[2]))
                        {
                            sStiffenerSize = curData[0].Trim('"').Trim();
                            break;
                        }
                    }
                }
            }

            if (String.IsNullOrEmpty(sStiffenerSize))
            {
                sStiffenerSize = "NOT AVAILABLE";
            }

            return sStiffenerSize;
        }

        private static decimal AdjustModulusToTemp(decimal mModulus, decimal mDuctTemp, bool bTempInCelcius)
        {
            if (!bTempInCelcius)
            {
                mDuctTemp = (mDuctTemp - 32M) * 5M / 9M;
            }

            return mModulus * (0.9883M - (0.00030913M * mDuctTemp) - (0.0000000013774M * UMCLib.Pow(mDuctTemp, 3M)));
        }

        private static bool ValidGauge(string sMaterial, int nGauge, bool bSpiral)
        {
            switch (sMaterial)
            {
                case "Steel":
                    if (!bSpiral)
                    {
                        switch (nGauge)
                        {
                            case 8:
                            case 10:
                            case 12:
                            case 14:
                            case 16:
                            case 18:
                            case 20:
                            case 22:
                            case 24:
                            case 26:
                            case 28:
                                return true;
                            default:
                                return false;
                        }
                    }
                    else
                    {
                        switch (nGauge)
                        {
                            case 14:
                            case 16:
                            case 18:
                            case 20:
                            case 22:
                            case 24:
                            case 26:
                            case 28:
                                return true;
                            default:
                                return false;
                        }
                    }
                case "Stainless Steel":
                    if (!bSpiral)
                    {
                        switch (nGauge)
                        {
                            case 8:
                            case 10:
                            case 12:
                            case 14:
                            case 16:
                            case 18:
                            case 20:
                            case 22:
                            case 24:
                            case 26:
                            case 28:
                                return true;
                            default:
                                return false;
                        }
                    }
                    else
                    {
                        switch (nGauge)
                        {
                            case 18:
                            case 20:
                            case 22:
                            case 24:
                            case 26:
                            case 28:
                                return true;
                            default:
                                return false;
                        }
                    }
                case "Aluminum":
                    if (!bSpiral)
                    {
                        switch (nGauge)
                        {
                            case 16:
                            case 18:
                            case 20:
                            case 22:
                            case 24:
                            case 26:
                            case 28:
                                return true;
                            default:
                                return false;
                        }
                    }
                    else
                    {
                        switch (nGauge)
                        {
                            case 20:
                            case 22:
                            case 24:
                            case 26:
                            case 28:
                                return true;
                            default:
                                return false;
                        }
                    }
                default:
                    return false;
            }
        }

        private static int ThicknessToGauge(string sMaterial, decimal mThickness)
        {
            int nGauge = 0;
            decimal mGaugeModulus = 0M;
            decimal mGaugeThickness = 0M;
            decimal mGaugeWeight = 0M;
            decimal mGaugeStrength = 0M;
            decimal mGaugeDensity = 0M;

            for (nGauge = 28; nGauge >= 8; nGauge -= 2)
            {
                try
                {
                    MetalProperties(sMaterial, nGauge, out mGaugeModulus, out mGaugeThickness, out mGaugeWeight, out mGaugeStrength, out mGaugeDensity);
                }
                catch { }

                if (mGaugeThickness >= mThickness)
                {
                    break;
                }
            }

            return nGauge;
        }

        private static void MetalProperties(string sMaterial, int nGauge, out decimal mModulus, out decimal mThickness, out decimal mUnitWeight, out decimal mYieldStrength, out decimal mDensity)
        {
            switch (sMaterial)
            {
                case "Steel":
                    mModulus = 30000000M;
                    mYieldStrength = 29000M;
                    mDensity = 0.2833M;
                    switch (nGauge)
                    {
                        case 28:
                            mThickness = 0.0157M;
                            mUnitWeight = 0.637M;
                            break;
                        case 26:
                            mThickness = 0.0187M;
                            mUnitWeight = .0759M;
                            break;
                        case 24:
                            mThickness = 0.0236M;
                            mUnitWeight = 0.959M;
                            break;
                        case 22:
                            mThickness = 0.0296M;
                            mUnitWeight = 1.204M;
                            break;
                        case 20:
                            mThickness = 0.0356M;
                            mUnitWeight = 1.449M;
                            break;
                        case 18:
                            mThickness = 0.0466M;
                            mUnitWeight = 1.897M;
                            break;
                        case 16:
                            mThickness = 0.0575M;
                            mUnitWeight = 2.342M;
                            break;
                        case 14:
                            mThickness = 0.0705M;
                            mUnitWeight = 2.873M;
                            break;
                        case 12:
                            mThickness = 0.0994M;
                            mUnitWeight = 4.052M;
                            break;
                        case 10:
                            mThickness = 0.1292M;
                            mUnitWeight = 5.268M;
                            break;
                        case 8:
                            mThickness = 0.1591M;
                            mUnitWeight = 6.487M;
                            break;
                        default:
                            throw new Exception("Unrecogined gauge. Valid gauges are 28, 26, 24, 22, 20, 18, 16, 14, 12, 10, 8");
                    }
                    break;
                case "Stainless Steel":
                    mModulus = 28000000M;
                    mYieldStrength = 30000M;
                    mDensity = 0.2917M;
                    switch (nGauge)
                    {
                        case 28:
                            mThickness = 0.0158M;
                            mUnitWeight = 0.656M;
                            break;
                        case 26:
                            mThickness = 0.0158M;
                            mUnitWeight = 0.788M;
                            break;
                        case 24:
                            mThickness = 0.0220M;
                            mUnitWeight = 1.050M;
                            break;
                        case 22:
                            mThickness = 0.0273M;
                            mUnitWeight = 1.313M;
                            break;
                        case 20:
                            mThickness = 0.0335M;
                            mUnitWeight = 1.575M;
                            break;
                        case 18:
                            mThickness = 0.0450M;
                            mUnitWeight = 2.100M;
                            break;
                        case 16:
                            mThickness = 0.0565M;
                            mUnitWeight = 2.625M;
                            break;
                        case 14:
                            mThickness = 0.0565M;
                            mUnitWeight = 3.281M;
                            break;
                        case 12:
                            mThickness = 0.0565M;
                            mUnitWeight = 4.594M;
                            break;
                        case 10:
                            mThickness = 0.0565M;
                            mUnitWeight = 5.906M;
                            break;
                        case 8:
                            mThickness = 0.0565M;
                            mUnitWeight = 7.219M;
                            break;
                        default:
                            throw new Exception("Unrecogined gauge. Valid gauges are 28, 26, 24, 22, 20, 18, 16, 14, 12, 10, 8");
                    }
                    break;
                case "Aluminum":
                    mModulus = 10000000M;
                    mYieldStrength = 20000M;
                    mDensity = 0.09999M;
                    switch (nGauge)
                    {
                        case 28:
                            mThickness = 0.0230M;
                            mUnitWeight = 0.3564M;
                            break;
                        case 26:
                            mThickness = 0.0230M;
                            mUnitWeight = 0.4562M;
                            break;
                        case 24:
                            mThickness = 0.0295M;
                            mUnitWeight = 0.5702M;
                            break;
                        case 22:
                            mThickness = 0.0365M;
                            mUnitWeight = 0.7128M;
                            break;
                        case 20:
                            mThickness = 0.0465M;
                            mUnitWeight = 0.898M;
                            break;
                        case 18:
                            mThickness = 0.0595M;
                            mUnitWeight = 1.141M;
                            break;
                        case 16:
                            mThickness = 0.0755M;
                            mUnitWeight = 1.283M;
                            break;
                        default:
                            throw new Exception("Unrecogined gauge. Valid gauges are 28, 26, 24, 22, 20, 18, 16");
                    }
                    break;
                default:
                    throw new Exception("Unrecogined material. Valid materials are Steel, Stainless Steel, Aluminum");
            }
        }

        private static decimal RoundCrossSectionalArea(decimal mDiameter)
        {
            return (3.14159M * UMCLib.Pow(mDiameter, 2M)) / 576M;
        }

        private static decimal CalcOrfPressDropAtVolFlow(decimal[] q, decimal[] p, decimal mSystemVolumeFlow)
        {
            decimal xTotal = 0;
            decimal yTotal = 0;
            decimal xyTotal = 0;
            decimal x2Total = 0;

            for (int i = 0; i < 7; i++)
            {
                decimal x = UMCLib.ConvertToDecimal(Math.Log10((double)p[i]));
                decimal y = UMCLib.ConvertToDecimal(Math.Log10((double)q[i]));
                decimal xy = x * y;
                decimal x2 = UMCLib.Pow(x, 2);

                xTotal += x;
                yTotal += y;
                xyTotal += xy;
                x2Total += x2;
            }

            decimal exponent = (7M * xyTotal - (xTotal * yTotal)) / (7M * x2Total - UMCLib.Pow(xTotal, 2M));
            decimal constant = UMCLib.Pow(10M, (yTotal - exponent * xTotal) / 7M);

            return UMCLib.Pow(mSystemVolumeFlow / constant, 1M / exponent);
        }

        private static decimal DuctVelocityCalc(decimal mVolumeFlowRate, decimal mArea)
        {
            return mVolumeFlowRate / mArea;
        }

        private static decimal PerfPressDrop(decimal mImpactVelocity, decimal mOpenArea)
        {
            decimal mPressureDrop1 = 0;
            decimal mPressureDrop2 = 0;
            decimal mPressureDrop3 = 0;
            decimal mPressureDrop4 = 0;
            decimal mPressureDrop5 = 0;
            decimal mPressureDrop6 = 0;
            decimal mPressureDrop7 = 0;
            decimal NUM = 0;
            decimal DEN = 0;

            if (mImpactVelocity >= 200M)
            {
                mPressureDrop1 = -0.00238M - 5.95238E-5M * mImpactVelocity + 1.07738E-5M * UMCLib.Pow(mImpactVelocity, 2M) + 4.166667E-9M * UMCLib.Pow(mImpactVelocity, 3);
                NUM = 1.3718E-11M + 6.25106E-9M * mImpactVelocity + 2.816566e-6M * UMCLib.Pow(mImpactVelocity, 2M);
                DEN = 1M - 3.1475373e-9M * mImpactVelocity - 7.3779467E-7M * UMCLib.Pow(mImpactVelocity, 2M) + 4.5018816E-10M * UMCLib.Pow(mImpactVelocity, 3M);
                mPressureDrop2 = NUM / DEN;
                NUM = -1.9164711E-7M - 1.9212053E-4M * mImpactVelocity + 1.8578987E-6M * UMCLib.Pow(mImpactVelocity, 2M);
                DEN = 1M - 4.4538352E-4M * mImpactVelocity + 1.899729e-7M * UMCLib.Pow(mImpactVelocity, 2M);
                mPressureDrop3 = NUM / DEN;
                mPressureDrop5 = 0.011138534M - 1.257935E-4M * mImpactVelocity + 9.7207405e-7M * UMCLib.Pow(mImpactVelocity, 2M);
                mPressureDrop4 = (mPressureDrop3 + mPressureDrop5) / 2M;
            }

            if (mImpactVelocity >= 500M)
            {
                mPressureDrop6 = -0.0007M - 3.3411954E-5M * mImpactVelocity + 2.03829E-7M * UMCLib.Pow(mImpactVelocity, 2M) - 1.030824E-12M * UMCLib.Pow(mImpactVelocity, 3M);
                NUM = 6.0888419E-5M + 4.587237E-5M * mImpactVelocity;
                DEN = 1M - 5.580479E-4M * mImpactVelocity + 1.031799827E-7M * mImpactVelocity;
                mPressureDrop7 = NUM / DEN;
            }

            if (mOpenArea >= 10 && mOpenArea < 15)
            {
                return ((mOpenArea - 10M) / (15M - 10M)) * (mPressureDrop2 - mPressureDrop1) + mPressureDrop1;
            }
            else if (mOpenArea >= 15 && mOpenArea < 20)
            {
                return ((mOpenArea - 15M) / (20M - 15M)) * (mPressureDrop3 - mPressureDrop2) + mPressureDrop2;
            }
            else if (mOpenArea >= 20 && mOpenArea < 25)
            {
                return ((mOpenArea - 20M) / (25M - 20M)) * (mPressureDrop4 - mPressureDrop3) + mPressureDrop3;
            }
            else if (mOpenArea >= 25 && mOpenArea < 30)
            {
                return ((mOpenArea - 25M) / (30M - 25M)) * (mPressureDrop5 - mPressureDrop4) + mPressureDrop4;
            }
            else if (mOpenArea >= 30 && mOpenArea < 50)
            {
                return ((mOpenArea - 30M) / (50M - 30M)) * (mPressureDrop6 - mPressureDrop5) + mPressureDrop5;
            }
            else if (mOpenArea >= 50 && mOpenArea < 63)
            {
                return ((mOpenArea - 50M) / (63M - 50M)) * (mPressureDrop7 - mPressureDrop6) + mPressureDrop6;
            }

            return 0M;
        }
    }
}
