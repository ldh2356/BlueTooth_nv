using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using InTheHand.Net;
using InTheHand.Net.Bluetooth;
using InTheHand.Net.Sockets;
using System.Threading;
using InTheHand.Net.Bluetooth.AttributeIds;

using System.Windows.Forms;
using System.Diagnostics;
using System.ComponentModel;

namespace DeviceAgents
{
    /// <summary>
    /// 에이전트 핸들러
    /// </summary>
    /// <param name="data"></param>
    public delegate void DeviceDataHandler(Bt32FeetDevice bt32FeetDevice, string data);

    /// <summary>
    /// 블루투스 컨트롤 클래스
    /// </summary>
    public class Bt32FeetDevice : IDisposable
    {
        ConnectLog log = new ConnectLog();

        /// <summary>
        /// 핸들러
        /// </summary>
        public DeviceDataHandler OnData = null;

        /// <summary>
        /// 에이전트로부터 받아올 맥 어드레스
        /// </summary>
        static string MacAddress { get; set; }


        /// <summary>
        /// 블루투스 어드레스 변수
        /// </summary>
        static BluetoothAddress bluetoothAddressString;


        /// <summary>
        /// 블루투스 디바이스 정보 클래스
        /// </summary>
        static BluetoothDeviceInfo bluetoothDeviceInfo;

        /// <summary>
        /// 워커 쓰레드
        /// </summary>
        protected Thread Worker = null;


        private int _lockcount;
        /// <summary>
        /// 락 카운트
        /// </summary>
        public int LockCount
        {
            get
            {
                return _lockcount;
            }
            set
            {
                _lockcount = value;
            }
        }

        /// <summary>
        /// IOS / Android 모델
        /// </summary>
        public int _model;


        /// <summary>
        /// 모바일 모델
        /// </summary>
        public EnumMobileModel mobileModel
        {
            get
            {
                if (_model == 0)
                    return EnumMobileModel.Android;
                else
                    return EnumMobileModel.IOS;
            }
        }


        /// <summary>
        /// UUID
        /// </summary>
        static string uuidStr = "00002415-0000-1000-8000-00805F9B34FB";

        /// <summary>
        /// GUID
        /// </summary>
        Guid uuid = new Guid(uuidStr);


        /// <summary>
        /// 생성자
        /// </summary>
        public Bt32FeetDevice()
        {
        }

        /// <summary>
        /// 쓰레드를 시작한다
        /// </summary>
        public virtual void Start()
        {
            try
            {
                Worker = new Thread(DoWork);
                Worker.Start();
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        /// 서비스를 정지한다
        /// </summary>
        public virtual void Stop()
        {
            try
            {
                if (Worker == null)
                    return;

                Worker.Interrupt();
                Worker = null;
            }
            catch (Exception)
            {

                throw;
            }
        }


        /// <summary>
        /// 주소를 져온다
        /// </summary>
        /// <param name="divAddr"></param>
        public void GetBtAddr(string divAddr)
        {
            MacAddress = divAddr;
        }


        /// <summary>
        /// Dispose 
        /// </summary>
        public void Dispose()
        {
            try
            {
                Stop();
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        /// 에이전트와 모바일 기기가 서로 통신이 되고있는지 여부 
        /// </summary>
        public bool IsServiced { get; set; } = true;



        /// <summary>
        /// 통신 체크
        /// </summary>
        protected void DoWork()
        {
            try
            {
                DoCheckBlueToothService();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        /// <summary>
        /// 블루투스 신호가 범위안에 잡히는지 테스트한다
        /// </summary>
        /// <returns></returns>
        private static bool IsInRange()
        {
            bool inRange = false;

            // 가짜 UUID 를 생성한다
            Guid fakeUuid = new Guid("{F13F471D-47CB-41d6-9609-BAD0690BF891}");
            BluetoothDeviceInfo device = new BluetoothDeviceInfo(BluetoothAddress.Parse(MacAddress));

            try
            {
                ServiceRecord[] records = device.GetServiceRecords(fakeUuid);
                inRange = true;
            }
            catch (Exception ex)
            {
                inRange = false;
            }

            return inRange;
        }

        /// <summary>
        /// 블루투스 서비스를 체크한다
        /// </summary>
        private void DoCheckBlueToothService()
        {
            try
            {
                // 파싱한 주소를 가져온다
                bluetoothAddressString = BluetoothAddress.Parse(MacAddress);

                // 컴퓨터에 블루투스가 연결되어있는지 여부를 확인 한다 
                while (true)
                {
                    // 블루투스 장치가 켜져있지 않다면 블루투스 설정 화면을 사용자에게 안내한다
                    if (!BluetoothRadio.IsSupported)
                    {
                        Process.Start("bthprops.cpl");
                        if (MessageBox.Show(Device.bluetoothOffMsg, "SSES", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                        {
                            break;
                        }
                    }
                    // 블루투스 제어장치가 확인돈다면
                    else
                    {
                        // 블루투스 제어장비를 초기화한다
                        bluetoothDeviceInfo = new BluetoothDeviceInfo(bluetoothAddressString);
                        break;
                    }
                }


                // 모바일 기기와 연결여부를 지속적으로 체크한다
                while (true)
                {
                    // 블루투스 라디오 신호가 범위안에 잡히는지 확인한다
                    if (IsInRange())
                    {
                        // 콜백 메서드로 이동
                        IAsyncResult iAsyncResult = bluetoothDeviceInfo.BeginGetServiceRecords(uuid, Service_AsyncCallback, bluetoothDeviceInfo);
                    }
                    // 잡히지 않는다면 Count 를 증가시킨다
                    else
                    {
                        log.write("==== 블루투스 라디오 범위가 넘어갔습니다 ====");
                        LockCount++;

                        // 아이폰의 경우 락카운트가 1 이상인경우 서비스 통신실패로 간주
                        if (EnumMobileModel.IOS == mobileModel)
                        {
                            log.write("==== IOS 락 카운트가 1 이상인 경우 서비스 통신 실패로 간주 ====");
                            if (LockCount > 1)
                            {
                                IsServiced = false;
                            }
                        }
                        // 안드로이드의 경우
                        else
                        {
                            log.write("==== 안드로이드 락 카운트가 3 이상인 경우 서비스 통신 실패로 간주 ====");
                            // 락 카운트가 3 이상인 경우 서비스 통신 실패로 간주
                            if (LockCount > 3)
                            {
                                IsServiced = false;
                            }
                        }
                    }

                    //이벤트 전달
                    if (OnData != null)
                        OnData(this, "");

                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                log.write("==== 통신 예외 발생 ====");
                log.write(ex.Message);
                Console.WriteLine(ex.Message + " Thread Exception!!! 1");
            }
        }



        /// <summary>
        /// 블루투스 통신 부분
        /// </summary>
        /// <param name="result"></param>
        private void Service_AsyncCallback(IAsyncResult iAsyncResult)
        {
            try
            {
                bluetoothDeviceInfo = iAsyncResult.AsyncState as BluetoothDeviceInfo;

                // 안드로이드의 경우
                if (mobileModel == EnumMobileModel.Android)
                {
                    try
                    {
                        ServiceRecord[] services = bluetoothDeviceInfo.EndGetServiceRecords(iAsyncResult);

                        // 안드로이드 기기와 통신이 성공한경우 (받아온 서비스가 0 개 이상일때)
                        if (services.Count() > 0)
                        {
                            IsServiced = true;
                            LockCount = 0;
                        }
                        // 실패한 경우 락카운트를 증가 시킨다
                        else
                        {
                            LockCount++;

                            // 락 카운트가 3 이상인 경우 서비스 통신 실패로 간주
                            if (LockCount > 3)
                            {
                                log.write("==== 락 카운트가 3 이상인 경우 서비스 통신 실패로 간주 ====");
                                IsServiced = false;
                            }
                        }
                    }
                    // 예외 발생시
                    catch (Exception ex)
                    {
                        LockCount++;


                        // 락 카운트가 3 이상인 경우 서비스 통신 실패로 간주
                        if (LockCount > 3)
                        {
                            log.write("==== 안드로이드 통신예외 ====");

                            IsServiced = false;
                        }
                    }
                }
                // IOS 의 경우
                else
                {
                    try
                    {
                        // 서비스를 가져올때 예외가 발생하지 않는다면 통신 성공으로 간주한다
                        ServiceRecord[] services = bluetoothDeviceInfo.EndGetServiceRecords(iAsyncResult);
                        IsServiced = true;
                        LockCount = 0;
                    }
                    // 실패한 경우 락카운트를 증가 시킨다
                    catch (Exception)
                    {
                        LockCount++;
                        // 락 카운트가 1 이상인 경우 서비스 통신 실패로 간주
                        if (LockCount > 1)
                        {
                            log.write("==== IOS 락 카운트가 3 이상인 경우 서비스 통신 실패로 간주 ====");
                            IsServiced = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString() + DateTime.Now.ToString());
            }
        }
    }
}
