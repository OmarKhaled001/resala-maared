
<!doctype html>
<html lang="{{ str_replace('_', '-', app()->getLocale()) }}" data-layout="vertical" data-sidebar="dark" data-sidebar-size="lg" data-preloader="disable" data-theme="default" data-bs-theme="light" data-topbar="light">
<title>الرئيسية</title>
<meta content="Minimal Admin & dashboard Template" name="description">
    @include('main.meta')
</head>
<body>
    <div id="layout-wrapper">
        @include('main.topbar')
        @include('main.sidebar')
        <div class="main-content">
            <div class="page-content">
                <div class="container-fluid">
                    <div class="row">
                        <div class="col-lg-12">
                          
                            <div class="row row-cols-xxl-5 row-cols-lg-3 row-cols-md-2 row-cols-1">
                                <div class="col">
                                    <div class="card border-bottom border-3 card-animate border-primary">
                                        <div class="card-body">
                                            <span class="badge bg-success-subtle text-success float-end">{{$masaolVolunteers->count()}} + {{$mmasaolVolunteers->count()}}</span>
                                            <h4 class="mb-4"><span class="counter-value" data-target="{{$masaolVolunteers->count() + $mmasaolVolunteers->count()}}">0</span></h4>
                            
                                            <p class="text-muted fw-medium text-uppercase mb-0">اجمالي فريق العمل</p>
                                        </div>
                                    </div>
                                </div><!--end col-->
                          
                                <div class="col">
                                    <div class="card border-bottom border-3 card-animate border-success">
                                        <div class="card-body">
                                            <span class="badge bg-success-subtle text-success float-end">{{ $mtotalContributionsCount}} / {{$masaolVolunteers->count()}}</span>
                                            <h4 class="mb-4"><span class="counter-value" data-target="{{round($mtotalSum/$masaolVolunteers->count(), 2) }}">0</span> /8</h4>
                                            <p class="text-muted fw-medium text-uppercase mb-0">متوسط مشاركات مسؤول</p>
                                        </div>
                                    </div>
                                </div><!--end col-->
                                <div class="col">
                                    <div class="card border-bottom border-3 card-animate  {{$mmtotalContributionsCount < $mmasaolVolunteers->count() ? 'border-danger' : 'border-success'}} ">
                                        <div class="card-body">
                                            <span class="badge {{$mmtotalContributionsCount < $mmasaolVolunteers->count() ? 'bg-danger-subtle text-danger' : 'bg-success-subtle text-success'}}  float-end">{{ $mmtotalContributionsCount}} / {{$mmasaolVolunteers->count()}}</span>
                                            <h4 class="mb-4"><span class="counter-value" data-target="{{round($mmtotalSum/$mmasaolVolunteers->count(), 2) }}">0</span> /8</h4>
                                            <p class="text-muted fw-medium text-uppercase mb-0">متوسط مشاركات مشروع مسؤول</p>
                                        </div>
                                    </div>
                                </div><!--end col-->
                              
                                <div class="col">
                                    <div class="card border-bottom border-3 card-animate border-warning">
                                        <div class="card-body">
                                            <span class="badge bg-success-subtle text-success float-end"><i class="ph-trend-up align-middle me-1"></i> 12.6%</span>
                                            <h4 class="mb-4"><span class="counter-value" data-target="4620">4,620</span></h4>
                            
                                            <p class="text-muted fw-medium text-uppercase mb-0">اجمالي المشاركات بدون تكرار</p>
                                        </div>
                                    </div>
                                </div><!--end col-->
                            
                                <div class="col">
                                    <div class="card border-bottom border-3 card-animate border-success">
                                        <div class="card-body">
                                            <span class="badge bg-success-subtle text-success float-end">{{$mmasaolVolunteers->count()}}</span>
                                            <h4 class="mb-4"><span class="counter-value" data-target="8541">8,541</span></h4>
                            
                                            <p class="text-muted fw-medium text-uppercase mb-0">متوسط مشاركات مشروع مسؤول</p>
                                        </div>
                                    </div>
                                </div><!--end col-->
                           
                        </div>
                    
                    </div>
                </div>
            </div>
            @include('main.footer')
        </div>
    </div>
    @include('main.scripts')
</body>
</html>
