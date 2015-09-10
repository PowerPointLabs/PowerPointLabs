namespace PowerPointLabs.ImageSearch.Domain
{
    class StyleOptionsFactory
    {
        public static StyleOptions GetOptions1()
        {
            var opt = new StyleOptions
            {
                IsDefaultOptions = false,
                FontColor = "#000000",
                // white circle banner
                BannerShape = 1,
                BannerOverlayColor = "#FFFFFF",
                BannerTransparency = 0,
                // white circle outline
                OutlineOverlayColor = "#FFFFFF",
                OutlineShape = 1,
                OutlineTransparency = 25,
                TextBoxPosition = 5,
                TextBoxOverlayColor = "#AAAAAA",
                TextBoxTransparency = 25,
                SpecialEffect = 4
            };
            return opt;
        }

        public static StyleOptions GetOptions2()
        {
            var opt = new StyleOptions
            {
                IsDefaultOptions = false,
                FontColor = "#FFFFFF",
                TextBoxPosition = 5,
                // black overlay
                OverlayColor = "#000000",
                Transparency = 25,
                // white rect banner
                BannerOverlayColor = "#D74926",
                BannerTransparency = 0,
                BannerDirection = 1,
                // white circle outline
                OutlineOverlayColor = "#D74926",
                OutlineTransparency = 25,
                TextBoxOverlayColor = "#000000",
                TextBoxTransparency = 85,
                SpecialEffect = 9
            };
            return opt;
        }

        public static StyleOptions GetOptions3()
        {
            return new StyleOptions { IsDefaultOptions = false };
        }

        public static StyleOptions GetOptions4()
        {
            return new StyleOptions { IsDefaultOptions = false };
        }

        public static StyleOptions GetOptions5()
        {
            return new StyleOptions { IsDefaultOptions = false };
        }
    }
}
