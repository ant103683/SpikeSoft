using SpikeSoft.UiUtils;
using SpikeSoft.UtilityManager;
using SpikeSoft.UtilityManager.TaskProgress;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SpikeSoft.ZS3Utilities.Tools.Meteor
{
    public class CharacterDataExport : GenericToolMenuItem<CharacterDataExport>
    {
        public CharacterDataExport() : base("Export Character Data to Excel") { }
        private List<string> messages = new List<string>()
        {
            "Select Character PAK Folder",
            "Select Save Folder",
            "Generating Excel data, Please Wait"
        };

        protected override void OnToolBtnClick(object sender, EventArgs e)
        {
            var paths = new List<string>();
            var charaDir = FileMan.GetDirectoryPath(messages[0]);
            if (charaDir == string.Empty)
            {
                return;
            }

            paths.Add(charaDir);

            var dir = new SaveFileDialog()
            {
                Title = messages[1],
                Filter = "xlsx|*.xlsx",
                FileName = "Character Param Info"
            };

            if (dir.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            paths.Add(dir.FileName);

            if (paths.Count() > 0) task(paths);
        }

        public async void task(List<string> data)
        {
            FunMan FUN = new FunMan();
            await FUN.InitializeTask(messages[2], new Action<object[], IProgress<ProgressInfo>>(work), new object[] { data }, false);
        }

        public void work(object[] args, IProgress<ProgressInfo> progress)
        {
            var paths = args[0] as List<string>;
            var charaDir = paths[0];
            var savePath = paths[1];
            var pakFiles = new List<string>();

            // Filter Character PAK Files from Directory
            // Only include files that match the pattern "_1p.pak"
            // Exclude files "Pilaf_parts_1p.pak" and less than 500000 bytes

            try
            {
                pakFiles = Directory.GetFiles(charaDir, "*_1p.pak")
                    .Where(file => !file.EndsWith("Pilaf_parts_1p.pak") && new FileInfo(file).Length >= 460000)
                    .ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while filtering PAK files: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Initialize Struct List
            List<CharacterBlastInfo> blasts = new List<CharacterBlastInfo>();
            List<CharacterSkillInfo> skills = new List<CharacterSkillInfo>();
            List<string> StructSource = new List<string>();
            var itemCounter = 0;


            try
            {
                foreach (var item in pakFiles)
                {
                    progress?.Report(new ProgressInfo { Value = (int)((itemCounter / (float)pakFiles.Count) * 100), Message = "Parsing PAK files..." });
                    int blastOffset = BinMan.GetBinaryData<int>(item, 96);
                    int skillOffset = BinMan.GetBinaryData<int>(item, 100);
                    var blastInfo = new StructMan<CharacterBlastInfo>(item, blastOffset);
                    blasts.Add(blastInfo[0]);
                    var skillInfo = new StructMan<CharacterSkillInfo>(item, skillOffset);
                    skills.Add(skillInfo[0]);
                    StructSource.Add(Path.GetFileName(item).Replace("_1p.pak", string.Empty));
                    itemCounter++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while processing PAK files: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Export to Excel
            var blastExcelFilename = Path.Combine(Path.GetDirectoryName(savePath), "CharacterBlastInfo.xlsx");
            var skillExcelFilename = Path.Combine(Path.GetDirectoryName(savePath), "CharacterSkillInfo.xlsx");
            XlsxMan xlsx1 = new XlsxMan(blastExcelFilename);
            xlsx1.ExportToExcel(StructSource, blasts);
            XlsxMan xlsx2 = new XlsxMan(skillExcelFilename);
            xlsx2.ExportToExcel(StructSource, skills);
        }
    }
}

public enum BlastMagic : short
{
    Teleport = 0,
    SolarFlare = 1,
    Paralysis = 2,
    TelekinesisChiaotzu = 3,
    Psychokinesis = 4,
    PsychokinesisKibitoKai = 5,
    Kaikosen7 = 7,
    Kaikosen8 = 8,
    ViceShoutGotenks = 9,
    ViceShoutBuu = 10,
    Acid = 11,
    AfterImage = 12,
    AndroidBarrier17 = 13,
    PsychoBarrierCooler = 14,
    PsychoBarrierBojack = 15,
    ExplosiveWave16 = 16,
    ExplosiveWave17 = 17,
    ExplosiveWave18 = 18,
    ExplosiveWaveA16 = 19,
    ExplosiveWave20 = 20,
    ExplosiveWaveBrolly = 21,
    JeiceExplosiveWave = 22,
    SelfHarm = 23,
    Kaioken = 24,
    LongAwaitedFor100P = 25,
    FullPower26 = 26,
    FullPowerNamekian = 27,
    FullPower28 = 28,
    FullPowerBojack = 29,
    FullPowerFrieza = 30,
    FullPowerBaby = 31,
    InexperiencedPowerUp = 32,
    SPFightingPose1 = 33,
    SPFightingPose2 = 34,
    SPFightingPose3 = 35,
    SPFightingPose4 = 36,
    SPFightingPose5 = 37,
    JusticeFinishingPose = 38,
    FalseCourage = 39,
    Howl = 40,
    FinishSign = 41,
    PumpUp = 42,
    SuperUnyieldingSpirit = 43,
    JusticePose1 = 44,
    HiTension = 45,
    FullPowerCharge = 46,
    PowerUpLimit = 47,
    Sleep = 48,
    DarkEyes = 50,
    AfterimageStrike = 51,
    MysticBreathBuu = 52,
    MysticBreathDabura = 53,
    MysticBreath = 54,
    WildSenze = 55,
    AndroidBarrier18 = 56,
    AndroidBarrier13 = 57,
    HerosFlute = 58,
    Kakarot = 59,
    SaiyanSoul = 60,
    JusticeFinishingPose2 = 61,
    BurningHeart = 62,
    FruitOfMightTree = 63,
    MakyoStar = 64,
    GiveMeEnergy = 65,
    Senzu = 66,
    PilafBarrier = 67,
    ExplosiveWaveSuuShenron = 68,
    ImSuperVegeta = 69,
    BeatYouIn5s = 70,
    KingOfSaiyans = 72,
    NowImMad = 73,
    PrincesPride = 74,
    MajinsAwakening = 75,
    RebirthCooler = 76,
    WizardBarrier = 77,
    Ncha = 78,
    SealedLightBeam = 80,
    Unforgivable = 81,
    KingsDignity = 82,
    ImTheWorst = 83,
    ForTheVillagers = 84,
    AllOut = 85,
    ImGettingExcited = 86,
    ChampionStyle = 87,
    DemonEye = 88,
    Shockwave = 89,
    PsychoThread = 90,
    Excited = 91,
    Stalling = 92,
    Barrier = 93,
    EvilBarrier = 94,
    MaidensWill = 95,
    MaidensExcitement = 96,
    MinionsLatentEnergy = 97,
    MadWarrior = 98,
    RebirthBuu = 99,
    TelekinesisGeneralBlue = 100,
    Kamehameha = 300,
    SuperKamehameha = 301,
    SuperKamehamehaEvilBuu = 302,
    Kamehamehax10 = 304,
    BigBangKamehameha = 305,
    OriginalKamehameha = 306,
    SuperKamekameha = 307,
    GalickGun = 308,
    GalickGunBaby = 309,
    SuperGalickGun = 310,
    FinalFlash = 311,
    FinalFlashBaby = 312,
    Masenko = 313,
    ChouMakouhou = 314,
    ChouMakouhouGreatApe = 315,
    RecoomeEraserGun = 316,
    FlameShowerBreath = 317,
    EvilImpulse = 318,
    TriBeam = 319,
    HellsFlash = 320,
    ElegantBlaster = 321,
    SoldierEnergyWave1 = 322,
    SoldierEnergyWave2 = 323,
    FullPowerEnergyWave1 = 324,
    FullPowerEnergyWave2 = 325,
    FullPowerEnergyWaveGreen = 326,
    FullPowerEnergyWaveRed = 327,
    FullPowerEnergyWaveKidTrunks = 328,
    FullPowerEnergyWaveCui = 329,
    FullPowerEnergyWaveBlue = 330,
    Makankosappo = 331,
    Dodonpa = 332,
    DeathBeam = 333,
    BigBangAttack = 334,
    LightGrenade = 335,
    FinishBuster = 336,
    EraserCannon = 337,
    FullPowerEnergyBall = 338,
    BurningAttack = 339,
    GrandSmasher = 340,
    GenocideBlast = 341,
    ExpandingEnergyWave = 342,
    DestructoDisc = 343,
    FissureSlash = 344,
    LightingShowerRain = 345,
    FullPowerEnergyBlastVolley = 346,
    FullPowerEnergyBlastVolleyGreen = 348,
    FullPowerEnergyBlastVolleyBlue = 349,
    FullPowerEnergyBlastVolleyPurple = 350,
    KaiokenAttack = 351,
    DrainLifeCell = 352,
    WolfFangFist = 353,
    PresentBomb = 354,
    HighSpeedRush = 355,
    HighPowerRush = 356,
    PsychicRockThrow = 357,
    GigantRockThrow = 358,
    GalacticDonutsGotenks = 359,
    GalacticDonutsBuu = 360,
    SuperExplosiveWave = 361,
    SuperExplosiveWaveA13 = 362,
    SuperExplosiveWaveHildergarn = 363,
    OriginalDodonpa = 364,
    CraneStyleAssassinStrike = 365,
    Counterattack = 366,
    HellsStorm = 367,
    DeathImpact = 368,
    HyperTornado = 369,
    PunishingBlaster = 370,
    VictoryCannon = 371,
    FinalImpact = 372,
    BraveCannon = 373,
    MaximumFlasher = 374,
    EnergysLast = 375,
    FingerBeam = 376,
    SSDeadlyBomber = 377,
    ShootBlaster = 378,
    CrusherBall = 379,
    FullPowerEnergyBallBojack = 380,
    GalaxyDynamite = 381,
    GiruMissile = 382,
    GekiretsuMadan = 383,
    SuperEnergyWaveVolleyPurple = 384,
    SuperEnergyWaveVolleyBlue = 385,
    SuperEnergyWaveVolleyYellow = 386,
    FullPowerEnergyBarrageWavePurple = 387,
    FullPowerEnergyBarrageWaveYellow = 388,
    FullPowerEnergyBarrageWaveCui = 389,
    InfinityBullet = 390,
    ShootingStarArrow = 391,
    DeathSaucer = 392,
    BurningStorm = 393,
    TrapShooterGreen = 394,
    TrapShooterYellow = 395,
    BraveSlash = 396,
    RapidCannon = 397,
    RecoomeRenegadeBomber = 398,
    DieDieMissileBarrage = 399,
    ChouMakouhouBarragePurple = 400,
    ChouMakouhouBarrageBlue = 401,
    ScatterFingerBeam = 402,
    DestructoDiscA18 = 404,
    FingerBlitzBarrageBlue = 405,
    FingerBlitzBarragePurple = 406,
    GiganticBlaze = 407,
    DeathStorm = 408,
    BlazingStorm = 409,
    HellsImpact = 410,
    EvilFlame = 411,
    GiganticFlame = 412,
    IMightDieThisTime = 413,
    DarknessEyeBeamYellow = 414,
    BionicPunisher = 415,
    DarknessEyeBeamPurple = 416,
    ChouMakousen = 417,
    KillDriver = 418,
    KaBlamSlicer = 419,
    LockOnBuster = 420,
    AmazingImpact = 421,
    UltimateImpact = 422,
    GiganticHammer = 424,
    GalacticTyrant = 425,
    CrazyRush = 426,
    FlashAndKill = 427,
    SilentAssassin13 = 428,
    JusticeSlash = 429,
    DarknessIllusion = 430,
    DeathChaser = 431,
    DesperadoRush = 432,
    Assault = 434,
    SpiritBreakingCannon = 435,
    NovaStrike = 436,
    PunishingRush = 437,
    FinalRevenger = 438,
    BloodyDance = 439,
    BlazingBarragePalm = 441,
    WildPressure = 442,
    TurtleSchoolUltimateFist = 443,
    TurtleSchoolFourVirtues = 444,
    ChouMaretsugeki = 445,
    JusticeCountdown = 447,
    SpaceMachAttack = 448,
    ChocolateBeam = 449,
    DragonThunder = 450,
    BurningShoot = 451,
    FinalGalickCannon = 452,
    MysticCombination = 453,
    DoubleBuster = 454,
    Soumasen = 455,
    GatlingGun = 456,
    RollingSmash = 457,
    YourNameIsDrum = 458,
    HiddenBlade = 459,
    BigBangCrash = 460,
    KamehamehaKaiokenx20 = 462,
    SuperExplosiveWaveBlue = 463,
    FullPowerEnergyWaveYellow = 464,
    MasenkoFutureGohan = 465,
    SuperMasenko = 466,
    ExplosiveMadan = 467,
    SuperExplosiveMadan = 468,
    HeresaPresent = 469,
    PhotonFlash = 470,
    ExplosiveDemonWave = 471,
    FullPowerDeathBeam = 472,
    BurstAttack = 473,
    GekiretsuMadanFutureGohan = 474,
    ExecutionBeam = 475,
    EnergyBallet = 476,
    BarrageDeathBeam = 477,
    PhotonStrike = 478,
    BakuretsuMahouko = 479,
    ShuraGekiretsuken = 480,
    MeteorCombination = 482,
    MeteorSmashGokuMed = 483,
    MeteorSmashGokuEnd = 484,
    MeteorCrashGokuEnd = 485,
    MeteorCrashGokuGT = 486,
    BloodySmash = 487,
    EightersAnger = 488,
    BurningBreaker = 489,
    DemonForkRush = 490,
    DynamicMessEmPunch = 492,
    CrazyCombination = 494,
    ImaTopClassWarrior = 495,
    InnocenseRush = 496,
    YakonIsNext = 497,
    BeamSwordSlash = 498,
    Skewer = 499,
    GiganticBomber = 500,
    PerfectCombination = 501,
    DragonFistGTBase = 502,
    Woohoo = 503,
    ForkAttack = 504,
    IllShootYou = 505,
    Ping = 506,
    EighterAttack = 507,
    BurningTornado = 508,
    DodoriaHeadBreaker = 509,
    CrossArmAttack = 510,
    PuiPuiNiceShot = 511,
    VolcanoExplosion = 512,
    BlazingStormNappa = 513,
    StayAwayFromMe = 514,
    ThatWontWork = 515,
    BerserkerCrash = 516,
    MadBanquet = 517,
    JusticeRush2 = 518,
    SuperKamehamehaGohanTeenSS2 = 519,
    BigBangKamehamehax100 = 601,
    UltimateSuperKamehameha = 602,
    InstantKamehameha = 603,
    MAXPOWERKamehameha = 604,
    UltimateSuperKamekameha = 605,
    UltimateFinalFlash = 606,
    NeoTriBeam = 608,
    UltimateElegantBlaster = 609,
    GalacticBuster = 613,
    DeathBall = 614,
    GenkidamaGokuMed = 616,
    UltimateBigBangAttack = 617,
    UltimateFinishBuster = 618,
    ShockingDeathBall = 620,
    HellzoneGrenade = 621,
    SuperGhostGotenks = 622,
    SuperGhostBuu = 623,
    CrazyFingerBeam = 624,
    DrainLife19 = 625,
    DrainLife20 = 626,
    UltimateDrainLifeCell = 627,
    MonsterCrush = 629,
    DarknessSwordAttack = 630,
    ShiningSwordAttack = 631,
    DimensionSwordAttack = 632,
    HAILFrieza = 633,
    BurstRush = 634,
    FarewellMrTien = 635,
    SaibamenBomb = 636,
    SelfDestructDevice = 637,
    Sadistic18 = 638,
    ChargingUltraBuuBuuVolleyball = 639,
    BodyChange = 640,
    GuldoSpecial = 642,
    StardustBreaker = 643,
    RecoomeFightingBomber = 644,
    AngryExplosion = 645,
    PerfectBarrier = 646,
    FinalExplosion = 647,
    RevengeDeathBomber = 648,
    UltimateSuperExplosiveWaveYellow = 649,
    UltimateSuperExplosiveWaveBlue = 650,
    UltimateSuperExplosiveWavePink = 651,
    UltimateSuperExplosiveWaveGreen = 652,
    Supernova = 653,
    RevengeDeathBall = 654,
    RevengeDeathBallFinal = 655,
    UltimateFinalPlan = 656,
    VidelRush = 657,
    GekiretsuShinouhou = 658,
    ThunderFlash = 659,
    UltimateSuperGalickGun = 660,
    UltimateVictoryCannon = 661,
    FinalShineAttack = 662,
    FinalSpiritCannon = 663,
    ShinGekiretsuShinouhou = 664,
    DarknessBlaster = 665,
    UltimateSSDeadlyBomber = 666,
    OmegaBlaster = 668,
    UltimateGrandSmasher = 669,
    PlanetBurst = 670,
    MinusEnergyPowerBall = 671,
    FullPowerEnergyBallAppule = 673,
    FullPowerEnergyBallGarlic = 674,
    ChouMakouhouBarrage = 675,
    UltimateGiganticBlaze = 676,
    UltimateGiganticFlame = 678,
    Soukidan = 679,
    DeadZone = 680,
    MankokuKyoutenshou = 681,
    LightningArrow = 682,
    UltimateChouMakousen1 = 683,
    UltimateChouMakousen2 = 684,
    SkyZapper = 685,
    IllusionSmash = 686,
    MaidensRage = 687,
    FatherSonKamehameha = 688,
    BrosKamehameha = 689,
    TurtleSchoolTranquility = 690,
    DirtyFireworks = 691,
    SalzaBladeRush = 692,
    JusticeJudgment = 693,
    PowerOfDarkness = 694,
    HeatDomeAttack = 695,
    BraveSwordAttack = 696,
    MiracleKaBlamSlash = 697,
    MeteorBurst = 698,
    DragonFistGokuEnd = 700,
    BakuretsuRanma = 701,
    AhLordFrieza = 702,
    UltimateEraserCannon = 703,
    LifeRiskingBlow = 704,
    LaunchMissiles = 705,
    UltimateExplosiveDemonWave = 706,
    SuperDodonWave = 707,
    GreatPilafOperation = 708,
    AngryKamehameha = 709,
    YouWillDieByMyHand = 710,
    MaximumBuster = 712,
    PlanetGeyser = 713,
    FinalKamehameha = 714,
    UltimateChouMakouhou = 715,
    SolarKamehameha = 716,
    BabidisUltimatePower = 717,
    Kapa = 718,
    UltimateSuperKamehamehaFutureGohan = 719,
    DetroyThePlanet = 720,
    GenkidamaGokuGT = 721,
    GenkidamaGokuEarly = 722,
    Begone = 723,
    UltimateMakankosappo = 724,
    DevilmiteBeam = 725,
    UnforgivableCell = 726,
    CrossArmDive = 727,
    HystericSaiyanLady = 728,
    DragonFistGokuGTSS3 = 729,
    DragonFistGokuGTSS4 = 730,
    PlayingProWrestling = 731,
    ColdFamilyPower = 732,
    PurpleCometAttackJeice = 733,
    PurpleCometAttackBurter = 734,
    Penetrate = 735,
    SuperMarengeki = 736,
    GiganticSaiyanLady = 737,
    SSDeadlyHammer = 738,
    YouHurtGoku = 739,
    MisterBuuArrives = 740,
    MysticFlasher = 741,
    RocketEngineSpark = 742,
    GekiretsuRanbu = 743,
    GenkidamaGokuEnd = 744,
    SadisticDance = 745,
    OrgaBlaster = 746,
    LightningSwordSlash = 747,
    BlasterMeteor = 748,
    GigaMeteorStorm = 749,
    SaveGoku = 750,
    MajinBuuResurrection = 751
}

[StructLayout(LayoutKind.Sequential, Pack = 1, Size = 640)]
public struct CharacterBlastInfo
{
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public uint[] mainProperties;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public uint[] secondaryProperties;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public BlastMagic[] MAGIC;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public ushort[] clashEffort;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field04;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field05;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field06;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field07;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field08;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field09;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field10;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field11;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field12;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field13;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field14;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field15;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field16;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field17;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field18;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field19;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field20;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field21;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field22;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field23;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field24;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field25;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field26;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field27;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field28;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field29;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field30;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field31;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field32;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field33;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field34;
    public byte padding_210;
    public byte padding_211;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field36;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field37;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field38;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field39;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field40;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field41;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field42;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field43;
    public int padding_308;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field44;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field45;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field46;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field47;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field48;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field49;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field50;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field51;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field52;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field53;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field54;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field55;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field56;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field57;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field58;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field59;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field60;
    public byte padding_363;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field61;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field62;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public int[] field63;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public int[] field64;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public int[] field65;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public int[] field66;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public int[] field67;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field68;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field69;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field70;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field71;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field72;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field73;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public short[] field74;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public short[] field75;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field76;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field77;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field78;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field79;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public byte[] field80;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field81;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    public float[] field82;
}

[Flags]
public enum B1_mainProperties : uint
{
    None = 0,
    unk_1 = 1 << 0,
    unk_2 = 1 << 1,
    unk_3 = 1 << 2,
    unk_4 = 1 << 3,
    unk_5 = 1 << 4,
    unk_6 = 1 << 5,
    unk_7 = 1 << 6,
    unk_8 = 1 << 7,
    unk_9 = 1 << 8,
    unk_10 = 1 << 9,
    unk_11 = 1 << 10,
    unk_12 = 1 << 11,
    unk_13 = 1 << 12,
    unk_14 = 1 << 13,
    unk_15 = 1 << 14,
    unk_16 = 1 << 15,
    unk_17 = 1 << 16,
    unk_18 = 1 << 17,
    unk_19 = 1 << 18,
    unk_20 = 1 << 19,
    unk_21 = 1 << 20,
    unk_22 = 1 << 21,
    unk_23 = 1 << 22,
    unk_24 = 1 << 23,
    unk_25 = 1 << 24,
    unk_26 = 1 << 25,
    unk_27 = 1 << 26,
    unk_28 = 1 << 27,
    unk_29 = 1 << 28,
    unk_30 = 1 << 29,
    unk_31 = 1 << 30,
    unk_32 = (uint)1 << 31,
}

[Flags]
public enum B1_secondaryProperties : uint
{
    None = 0,
    OnlyInLockOn = 1 << 0,
    LookOpponent = 1 << 1,
    TurnAroundToOpponent = 1 << 2,
    UsableWhileReceivingDamage = 1 << 3,
    OneTime = 1 << 4,
    unk_6 = 1 << 5,
    unk_7 = 1 << 6,
    SetStatsInEnd = 1 << 7,
    EffectsStack = 1 << 8,
    MaxMode = 1 << 9,
    PowerArmor = 1 << 10,
    MeleeArmor = 1 << 11,
    unk_13 = 1 << 12,
    unk_14 = 1 << 13,
    unk_15 = 1 << 14,
    unk_16 = 1 << 15,
    unk_17 = 1 << 16,
    unk_18 = 1 << 17,
    CameraFocusOnUser = 1 << 18,
    unk_20 = 1 << 19,
    unk_21 = 1 << 20,
    Unblockable = 1 << 21,
    EvasionPerformable = 1 << 22,
    KiChargePenalty = 1 << 23,
    InvulnerabilityMelee = 1 << 24,
    InvulnerabilityRushBlast2 = 1 << 25,
    InvulnerabilityBlast2 = 1 << 26,
    InvulnerabilityStunMoves = 1 << 27,
    unk_29 = 1 << 28,
    unk_30 = 1 << 29,
    unk_31 = 1 << 30,
    unk_32 = (uint)1 << 31,
}

[StructLayout(LayoutKind.Sequential, Pack = 1, Size = 256)]
public struct CharacterSkillInfo
{
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public B1_mainProperties[] mainProperties;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public B1_secondaryProperties[] secondaryProperties;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public BlastMagic[] MAGIC;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public ushort[] clashEffort;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public float[] vfxDuration;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public float[] vfxSpeed;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public float[] vfxDispersion;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public float[] vfxSize;
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field08; // 0x38
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] hitDamageDivisorMultiplier; // 0x3A
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] energyHitDamageDivisorMultiplier; // 0x3C
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field10_b; // 0x3E
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] hitCount; // 0x40
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field12; // 0x42
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field13; // 0x44
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field14; // 0x46
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field15; // 0x48
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field16; // 0x4A
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field17; // 0x4C
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field18; // 0x4E
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field19; // 0x50
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field20; // 0x52
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field21; // 0x54
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field22; // 0x56
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field23; // 0x58
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field24; // 0x5A
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field25; // 0x5C
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field26; // 0x5E
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field27; // 0x60
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field28; // 0x62
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] animationType; // 0x64
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] statEffectLifespanType; // 0x66
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field31; // 0x68
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] secondaryHitEffect; // 0x6A
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field33_a; // 0x6C
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] blurInvocation; // 0x6E
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] blurFinish; // 0x70
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] field35; // 0x72
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public float[] knockback; // 0x74
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public float[] blockingKnockback; // 0x7C
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public int[] damage; // 0x84
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public int[] blockedDamage; // 0x8C
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] cost; // 0x94
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] attackAbility; // 0x96
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] kiAbility; // 0x98
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] defenseAbility; // 0x9A
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] superAbility; // 0x9C
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public byte[] comType; // 0x9E
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public float[] lifespan; // 0xA0
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public int[] hpModifier; // 0xA8
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public int[] kiModifier; // 0xB0
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public int[] field49; // 0xB8
    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
    public int[] field50; // 0xC0
}